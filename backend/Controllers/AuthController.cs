using System;
using System.IdentityModel.Tokens.Jwt;
using System.Linq;
using System.Security.Claims;
using System.Text;
using ConverterApi.Data;
using ConverterApi.Models;
using ConverterApi.Services;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Configuration;
using Microsoft.IdentityModel.Tokens;

namespace ConverterApi.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class AuthController : ControllerBase
    {
        private readonly AppDbContext _db;
        private readonly IConfiguration _config;
        private readonly IEmailService _emailService;

        public AuthController(AppDbContext db, IConfiguration config, IEmailService emailService)
        {
            _db = db;
            _config = config;
            _emailService = emailService;
        }

        [HttpPost("register")]
        public async System.Threading.Tasks.Task<IActionResult> Register([FromBody] RegisterRequest req)
        {
            if (string.IsNullOrWhiteSpace(req.Email) || string.IsNullOrWhiteSpace(req.Password))
                return BadRequest("Email ve şifre gerekli");
            
            if (string.IsNullOrWhiteSpace(req.FirstName) || string.IsNullOrWhiteSpace(req.LastName))
                return BadRequest("İsim ve soyisim gerekli");

            if (await _db.Users.AnyAsync(u => u.Email == req.Email))
                return Conflict("Email zaten kayıtlı");

            PasswordHasher.CreatePasswordHash(req.Password, out var hash, out var salt);
            
            var verificationToken = Convert.ToBase64String(Guid.NewGuid().ToByteArray()).Replace("/", "_").Replace("+", "-");
            
            var user = new User { 
                Email = req.Email, 
                PasswordHash = hash, 
                PasswordSalt = salt,
                FirstName = req.FirstName,
                LastName = req.LastName,
                PhoneCountryCode = req.CountryCode,
                PhoneNumber = req.PhoneNumber,
                VerificationToken = verificationToken,
                IsVerified = false
            };
            _db.Users.Add(user);
            await _db.SaveChangesAsync();

            // Send Verification Email (Non-blocking / Fire-and-forget)
            // Render Free Tier SMTP portlarını engellediği için mail beklemek UI'ı dondurur.
            var frontendUrl = (_config["FrontendUrl"] ?? _config["FRONTEND_URL"] ?? "https://file-converter-phi-nine.vercel.app").Trim();
            if (frontendUrl.EndsWith("/")) frontendUrl = frontendUrl.TrimEnd('/');
            
            var verifyLink = $"{frontendUrl}/verify.html?token={verificationToken}";
            
            _ = System.Threading.Tasks.Task.Run(async () => {
                try {
                    await _emailService.SendVerificationEmailAsync(user.Email, verifyLink);
                } catch (Exception ex) {
                    Console.WriteLine($"BACKGROUND EMAIL ERROR: {ex.Message}");
                }
            });
            
            // Console log for dev (optional, can be kept for server logs)
            Console.WriteLine($"VERIFICATION LINK: {verifyLink}");

            return Ok(new { message = "Kayıt başarılı. Lütfen emailinize gelen doğrulama linkine tıklayın.", verifyLink });
        }

        [HttpPost("login")]
        public async System.Threading.Tasks.Task<IActionResult> Login([FromBody] LoginRequest req)
        {
            var user = await _db.Users.SingleOrDefaultAsync(u => u.Email == req.Email);
            if (user == null) return Unauthorized("Geçersiz kimlik bilgileri");
            
            if (!user.IsVerified)
                return Unauthorized("Lütfen önce email adresinizi doğrulayın.");

            if (!PasswordHasher.VerifyPassword(req.Password, user.PasswordHash, user.PasswordSalt))
                return Unauthorized("Geçersiz kimlik bilgileri");

            var token = GenerateJwtToken(user);
            return Ok(new { token, user = new { user.Id, user.Email } });
        }

        [HttpPost("verify-email")]
        public async System.Threading.Tasks.Task<IActionResult> VerifyEmail([FromBody] VerifyEmailRequest req)
        {
            var user = await _db.Users.SingleOrDefaultAsync(u => u.VerificationToken == req.Token);
            if (user == null) return BadRequest("Geçersiz doğrulama kodu");

            user.IsVerified = true;
            user.VerificationToken = null; // Token'ı temizle
            await _db.SaveChangesAsync();

            return Ok(new { message = "Email başarıyla doğrulandı. Giriş yapabilirsiniz." });
        }

        [HttpPost("forgot-password")]
        public async System.Threading.Tasks.Task<IActionResult> ForgotPassword([FromBody] ForgotRequest req)
        {
            var user = await _db.Users.SingleOrDefaultAsync(u => u.Email == req.Email);
            if (user == null) return Ok(); // Security: don't reveal if email exists
            
            // Mevcut reset token'ları invalidate et
            var oldTokens = _db.PasswordResetTokens.Where(t => t.UserId == user.Id && !t.IsUsed);
            _db.PasswordResetTokens.RemoveRange(oldTokens);

            // Yeni reset token oluştur
            var resetToken = Convert.ToBase64String(Guid.NewGuid().ToByteArray()).Substring(0, 12);
            var resetTokenEntity = new PasswordResetToken
            {
                UserId = user.Id,
                Token = resetToken,
                ExpiresAt = DateTime.UtcNow.AddHours(1)
            };
            _db.PasswordResetTokens.Add(resetTokenEntity);
            await _db.SaveChangesAsync();

            // Email gönder (dev için token'ı da döndür)
            var frontendUrl = _config["FrontendUrl"] ?? "http://localhost:8000";
            var resetLink = $"{frontendUrl}?resetToken={resetToken}";
            await _emailService.SendPasswordResetEmailAsync(user.Email, resetToken, resetLink);

            return Ok(new { resetToken, message = "Sıfırlama kodu email'e gönderildi. (Dev: Kod da burada)" });
        }

        [HttpPost("reset-password")]
        public async System.Threading.Tasks.Task<IActionResult> ResetPassword([FromBody] ResetPasswordRequest req)
        {
            if (string.IsNullOrWhiteSpace(req.Token) || string.IsNullOrWhiteSpace(req.NewPassword))
                return BadRequest("Token ve yeni şifre gerekli");

            var resetToken = await _db.PasswordResetTokens
                .Include(t => t.User)
                .SingleOrDefaultAsync(t => t.Token == req.Token && !t.IsUsed && t.ExpiresAt > DateTime.UtcNow);

            if (resetToken == null)
                return BadRequest("Geçersiz veya süresi dolmuş token");

            var user = resetToken.User;
            if (user == null) return BadRequest("Kullanıcı bulunamadı");

            PasswordHasher.CreatePasswordHash(req.NewPassword, out var hash, out var salt);
            user.PasswordHash = hash;
            user.PasswordSalt = salt;

            resetToken.IsUsed = true;
            _db.Users.Update(user);
            _db.PasswordResetTokens.Update(resetToken);
            await _db.SaveChangesAsync();

            return Ok(new { message = "Şifre başarıyla değiştirildi" });
        }

        private string GenerateJwtToken(User user)
        {
            var key = Encoding.UTF8.GetBytes(_config["Jwt:Key"]!);
            var creds = new SigningCredentials(new SymmetricSecurityKey(key), SecurityAlgorithms.HmacSha256);
            var claims = new[] { new Claim(JwtRegisteredClaimNames.Sub, user.Email), new Claim("id", user.Id.ToString()) };
            var token = new JwtSecurityToken(
                issuer: _config["Jwt:Issuer"],
                audience: _config["Jwt:Audience"],
                claims: claims,
                expires: DateTime.UtcNow.AddMinutes(double.Parse(_config["Jwt:ExpireMinutes"]!)),
                signingCredentials: creds
            );
            return new JwtSecurityTokenHandler().WriteToken(token);
        }
    }

    public record RegisterRequest(string Email, string Password, string FirstName, string LastName, string? CountryCode, string? PhoneNumber);
    public record LoginRequest(string Email, string Password);
    public record ForgotRequest(string Email);
    public record ResetPasswordRequest(string Token, string NewPassword);
    public record VerifyEmailRequest(string Token);
}
