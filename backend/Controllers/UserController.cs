using System;
using System.Linq;
using System.Security.Claims;
using System.Threading.Tasks;
using ConverterApi.Data;
using ConverterApi.Models;
using ConverterApi.Services;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;

namespace ConverterApi.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    [Authorize]
    public class UserController : ControllerBase
    {
        private readonly AppDbContext _db;

        public UserController(AppDbContext db)
        {
            _db = db;
        }

        [HttpGet("profile")]
        public async Task<IActionResult> GetProfile()
        {
            var userId = GetCurrentUserId();
            if (userId == null) return Unauthorized();

            var user = await _db.Users.FindAsync(userId);
            if (user == null) return NotFound("Kullanıcı bulunamadı");

            return Ok(new
            {
                user.Id,
                user.Email,
                user.FirstName,
                user.LastName,
                user.PhoneNumber,
                user.CreatedAt
            });
        }

        [HttpPut("profile")]
        public async Task<IActionResult> UpdateProfile([FromBody] UpdateProfileRequest req)
        {
            var userId = GetCurrentUserId();
            if (userId == null) return Unauthorized();

            var user = await _db.Users.FindAsync(userId);
            if (user == null) return NotFound("Kullanıcı bulunamadı");

            user.FirstName = req.FirstName;
            user.LastName = req.LastName;
            user.PhoneNumber = req.PhoneNumber;

            _db.Users.Update(user);
            await _db.SaveChangesAsync();

            return Ok(new { message = "Profil güncellendi", user = new { user.Id, user.Email, user.FirstName, user.LastName, user.PhoneNumber } });
        }

        [HttpPost("change-password")]
        public async Task<IActionResult> ChangePassword([FromBody] ChangePasswordRequest req)
        {
            var userId = GetCurrentUserId();
            if (userId == null) return Unauthorized();

            if (string.IsNullOrWhiteSpace(req.OldPassword) || string.IsNullOrWhiteSpace(req.NewPassword))
                return BadRequest("Eski ve yeni şifre gerekli");

            var user = await _db.Users.FindAsync(userId);
            if (user == null) return NotFound("Kullanıcı bulunamadı");

            // Verify old password
            if (!PasswordHasher.VerifyPassword(req.OldPassword, user.PasswordHash, user.PasswordSalt))
                return BadRequest("Eski şifre yanlış");

            // Set new password
            PasswordHasher.CreatePasswordHash(req.NewPassword, out var hash, out var salt);
            user.PasswordHash = hash;
            user.PasswordSalt = salt;

            _db.Users.Update(user);
            await _db.SaveChangesAsync();

            return Ok(new { message = "Şifre başarıyla değiştirildi" });
        }

        private int? GetCurrentUserId()
        {
            var userIdClaim = User.FindFirst("id")?.Value;
            if (userIdClaim == null) return null;
            return int.TryParse(userIdClaim, out var userId) ? userId : null;
        }
    }

    public record UpdateProfileRequest(string? FirstName, string? LastName, string? PhoneNumber);
    public record ChangePasswordRequest(string OldPassword, string NewPassword);
}
