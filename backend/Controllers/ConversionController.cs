using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using PdfSharpCore.Drawing;
using PdfSharpCore.Pdf;
using System.IO;
using ConverterApi.Data;
using ConverterApi.Models;
using ConverterApi.Services;
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
        private readonly IConversionService _conversionService;
        private readonly ILogger<ConversionController> _logger;

        public ConversionController(
            AppDbContext context, 
            IWebHostEnvironment environment, 
            IConversionService conversionService,
            ILogger<ConversionController> logger)
        {
            _context = context;
            _environment = environment;
            _conversionService = conversionService;
            _logger = logger;
        }

        [AllowAnonymous]
        [HttpPost("convert")]
        public async Task<IActionResult> ConvertFile([FromForm] IFormFile file, [FromForm] string sourceFormat, [FromForm] string targetFormat)
        {
            if (file == null || file.Length == 0)
                return BadRequest("Dosya yüklenmedi.");

            var source = (sourceFormat ?? Path.GetExtension(file.FileName).TrimStart('.')).ToLower();
            var target = targetFormat.ToLower();

            try
            {
                byte[] resultBytes;
                string contentType;
                string downloadFileName;

                // Mevcut Resim -> PDF mantığını koruyoruz (PdfSharpCore ile)
                if ((source == "jpg" || source == "jpeg" || source == "png") && target == "pdf")
                {
                    resultBytes = await ImageToPdfAsync(file);
                    contentType = "application/pdf";
                    downloadFileName = Path.GetFileNameWithoutExtension(file.FileName) + ".pdf";
                }
                else
                {
                    // Diğer tüm dönüşümler için yeni servisi kullanıyoruz
                    using var stream = file.OpenReadStream();
                    resultBytes = await _conversionService.ConvertAsync(stream, file.FileName, source, target);
                    
                    var extension = GetTargetExtension(target, resultBytes);
                    contentType = GetContentType(extension);
                    downloadFileName = Path.GetFileNameWithoutExtension(file.FileName) + extension;
                }

                // Geçmişe kaydet (yalnızca giriş yapmış kullanıcılar için)
                var userIdClaim = User.FindFirst("id");
                if (userIdClaim != null)
                {
                    await SaveToHistoryAsync(int.Parse(userIdClaim.Value), file.FileName, target, resultBytes);
                }

                return File(resultBytes, contentType, downloadFileName);
            }
            catch (NotSupportedException ex)
            {
                return BadRequest(ex.Message);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Dönüştürme hatası: {FileName}", file.FileName);
                return StatusCode(500, $"Dönüştürme Hatası: {ex.Message}");
            }
        }

        private async Task<byte[]> ImageToPdfAsync(IFormFile file)
        {
            using var ms = new MemoryStream();
            await file.CopyToAsync(ms);
            ms.Position = 0;

            using var pdf = new PdfDocument();
            var page = pdf.AddPage();
            using (var image = XImage.FromStream(() => new MemoryStream(ms.ToArray())))
            {
                page.Width = image.PointWidth;
                page.Height = image.PointHeight;
                using (var gfx = XGraphics.FromPdfPage(page))
                {
                    gfx.DrawImage(image, 0, 0, image.PointWidth, image.PointHeight);
                }
            }
            using var output = new MemoryStream();
            pdf.Save(output);
            return output.ToArray();
        }

        private async Task SaveToHistoryAsync(int userId, string originalFileName, string targetFormat, byte[] resultBytes)
        {
            var webRoot = _environment.WebRootPath ?? Path.Combine(Directory.GetCurrentDirectory(), "wwwroot");
            var uploadDir = Path.Combine(webRoot, "files", userId.ToString());
            if (!Directory.Exists(uploadDir)) Directory.CreateDirectory(uploadDir);

            var extension = GetTargetExtension(targetFormat, resultBytes);
            var outputFileName = $"{Guid.NewGuid()}{extension}";
            var outputPath = Path.Combine(uploadDir, outputFileName);

            await System.IO.File.WriteAllBytesAsync(outputPath, resultBytes);

            var record = new ConversionRecord
            {
                UserId = userId,
                OriginalFileName = originalFileName,
                ConvertedFileName = outputFileName,
                FilePath = outputPath,
                FileSize = resultBytes.Length,
                CreatedAt = DateTime.UtcNow
            };

            _context.ConversionRecords.Add(record);
            await _context.SaveChangesAsync();
        }

        private string GetTargetExtension(string target, byte[] bytes)
        {
            // PDF -> Image (Multiple pages) returns ZIP
            if ((target == "png" || target == "jpg" || target == "jpeg") && IsZip(bytes))
                return ".zip";

            return target.ToLower() switch
            {
                "pdf" => ".pdf",
                "docx" => ".docx",
                "xlsx" => ".xlsx",
                "pptx" => ".pptx",
                "png" => ".png",
                "jpg" or "jpeg" => ".jpg",
                _ => "." + target
            };
        }

        private bool IsZip(byte[] bytes) => bytes.Length > 4 && bytes[0] == 0x50 && bytes[1] == 0x4B && bytes[2] == 0x03 && bytes[3] == 0x04;

        private string GetContentType(string extension)
        {
            return extension.ToLower() switch
            {
                ".pdf" => "application/pdf",
                ".docx" => "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                ".xlsx" => "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                ".pptx" => "application/vnd.openxmlformats-officedocument.presentationml.presentation",
                ".png" => "image/png",
                ".jpg" or ".jpeg" => "image/jpeg",
                ".zip" => "application/zip",
                _ => "application/octet-stream"
            };
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

            if (record == null) return NotFound("Dosya bulunamadı.");
            if (!System.IO.File.Exists(record.FilePath)) return NotFound("Dosya fiziksel olarak silinmiş.");

            var bytes = await System.IO.File.ReadAllBytesAsync(record.FilePath);
            var extension = Path.GetExtension(record.FilePath);
            return File(bytes, GetContentType(extension), record.OriginalFileName.Replace(Path.GetExtension(record.OriginalFileName), extension));
        }
    }
}
