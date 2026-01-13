using ConverterApi.Data;
using ConverterApi.Models;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using System.Security.Claims;

namespace ConverterApi.Controllers
{
    [Authorize]
    [ApiController]
    [Route("api/[controller]")]
    public class SupportController : ControllerBase
    {
        private readonly AppDbContext _context;

        public SupportController(AppDbContext context)
        {
            _context = context;
        }

        [AllowAnonymous]
        [HttpPost("submit")]
        public async Task<IActionResult> SubmitRequest([FromBody] SupportRequestDto request)
        {
            try
            {
                var userIdClaim = User.FindFirst("id");
                int? userId = userIdClaim != null ? int.Parse(userIdClaim.Value) : null;

                var supportRequest = new SupportRequest
                {
                    UserId = userId,
                    Subject = request.Subject,
                    Message = request.Message,
                    CreatedAt = DateTime.UtcNow
                };

                _context.SupportRequests.Add(supportRequest);
                await _context.SaveChangesAsync();

                return Ok(new { message = "Talebiniz başarıyla alındı. Teşekkürler!" });
            }
            catch (Exception ex)
            {
                // Detailed error for debugging
                return StatusCode(500, new { 
                    message = "Talep gönderilirken sunucu hatası oluştu.", 
                    error = ex.Message,
                    details = ex.InnerException?.Message 
                });
            }
        }
    }

    public class SupportRequestDto
    {
        public string Subject { get; set; } = string.Empty;
        public string Message { get; set; } = string.Empty;
    }
}
