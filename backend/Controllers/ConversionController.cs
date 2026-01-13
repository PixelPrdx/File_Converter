using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using PdfSharpCore.Drawing;
using PdfSharpCore.Pdf;
using PdfSharpCore.Drawing;
using PdfSharpCore.Pdf;
using System.IO;
using ConverterApi.Data;
using ConverterApi.Models;
using System.Security.Claims;
using Microsoft.EntityFrameworkCore;

namespace ConverterApi.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    [Authorize]
    public class ConversionController : ControllerBase
    {
        private readonly AppDbContext _context;
        private readonly IWebHostEnvironment _environment;

        public ConversionController(AppDbContext context, IWebHostEnvironment environment)
        {
            _context = context;
            _environment = environment;
        }

        [AllowAnonymous]
        [HttpPost("convert")]
        public async Task<IActionResult> ConvertFile([FromForm] IFormFile file, [FromForm] string sourceFormat, [FromForm] string targetFormat)
        {
            if (file == null || file.Length == 0)
                return BadRequest("Dosya yüklenmedi.");

            if (string.IsNullOrEmpty(targetFormat) || targetFormat.ToLower() != "pdf")
                return BadRequest("Şu an sadece PDF hedef formatı desteklenmektedir.");

            // Basic validation
            var extension = Path.GetExtension(file.FileName).ToLower().TrimStart('.');
            if (extension != sourceFormat.ToLower())
            {
                // Strict check allows only if extension matches selection. 
                // Alternatively, just checking if supported image types.
                var supportedImages = new[] { "jpg", "jpeg", "png" };
                if (!supportedImages.Contains(extension))
                    return BadRequest("Desteklenmeyen dosya formatı. Sadece JPG ve PNG destekleniyor.");
            }

            try
            {
                using (var memoryStream = new MemoryStream())
                {
                    file.CopyTo(memoryStream);
                    memoryStream.Position = 0;

                    using (var pdfDocument = new PdfDocument())
                    {
                        var page = pdfDocument.AddPage();
                        
                        // Load image
                        using (var image = XImage.FromStream(() => new MemoryStream(memoryStream.ToArray())))
                        {
                            page.Width = image.PointWidth;
                            page.Height = image.PointHeight;

                            using (var gfx = XGraphics.FromPdfPage(page))
                            {
                                gfx.DrawImage(image, 0, 0, image.PointWidth, image.PointHeight);
                            }
                        }

                            
                            // Handle User ID for history
                            var userIdClaim = User.FindFirst("id");
                            int? userId = userIdClaim != null ? int.Parse(userIdClaim.Value) : null;
                            
                            // Robust path handling
                            var webRoot = _environment.WebRootPath;
                            if (string.IsNullOrEmpty(webRoot))
                            {
                                webRoot = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot");
                            }

                            // Use "guest" or userId for folder
                            var userFolder = userId?.ToString() ?? "guest";
                            var uploadDir = Path.Combine(webRoot, "files", userFolder);
                            if (!Directory.Exists(uploadDir))
                                Directory.CreateDirectory(uploadDir);

                            var outputFileName = $"{Guid.NewGuid()}.pdf";
                            var outputPath = Path.Combine(uploadDir, outputFileName);
                            
                            // Save PDF to disk
                            pdfDocument.Save(outputPath);

                            // Save history record ONLY for logged-in users
                            if (userId.HasValue)
                            {
                                var record = new ConversionRecord
                                {
                                    UserId = userId.Value,
                                    OriginalFileName = file.FileName,
                                    ConvertedFileName = outputFileName,
                                    FilePath = outputPath,
                                    FileSize = new FileInfo(outputPath).Length,
                                    CreatedAt = DateTime.UtcNow
                                };

                                _context.ConversionRecords.Add(record);
                                await _context.SaveChangesAsync();
                            }

                            // Return file for immediate download
                            var pdfBytes = await System.IO.File.ReadAllBytesAsync(outputPath);
                            var displayFileName = Path.GetFileNameWithoutExtension(file.FileName) + ".pdf";
                            
                            return File(pdfBytes, "application/pdf", displayFileName);
                        }
                    }
                }
            catch (Exception ex)
            {
                Console.WriteLine($"Conversion Error: {ex}");
                return StatusCode(500, $"Dönüştürme Hatası: {ex.Message}");
            }
        }

        [HttpGet("history")]
        public async Task<IActionResult> GetHistory()
        {
            var userId = int.Parse(User.FindFirst("id")!.Value);
            var history = await _context.ConversionRecords
                .Where(c => c.UserId == userId)
                .OrderByDescending(c => c.CreatedAt)
                .ToListAsync();

            return Ok(history);
        }

        [HttpGet("download/{id}")]
        public async Task<IActionResult> DownloadFile(int id)
        {
            var userId = int.Parse(User.FindFirst("id")!.Value);
            var record = await _context.ConversionRecords.FirstOrDefaultAsync(c => c.Id == id && c.UserId == userId);

            if (record == null)
                return NotFound("Dosya bulunamadı.");

            if (!System.IO.File.Exists(record.FilePath))
                return NotFound("Dosya fiziksel olarak silinmiş.");

            var bytes = await System.IO.File.ReadAllBytesAsync(record.FilePath);
            var fileName = Path.GetFileNameWithoutExtension(record.OriginalFileName) + ".pdf";

            return File(bytes, "application/pdf", fileName);
        }
    }
}
