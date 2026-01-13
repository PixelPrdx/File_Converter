using System;
using System.ComponentModel.DataAnnotations;

namespace ConverterApi.Models
{
    public class User
    {
        public int Id { get; set; }

        [Required]
        [MaxLength(256)]
        public string Email { get; set; } = null!;

        [MaxLength(100)]
        public string? FirstName { get; set; }

        [MaxLength(100)]
        public string? LastName { get; set; }

        [MaxLength(8)]
        public string? PhoneCountryCode { get; set; }

        [MaxLength(32)]
        public string? PhoneNumber { get; set; }

        [MaxLength(100)]
        public string? VerificationToken { get; set; }
        
        public bool IsVerified { get; set; } = false;

        [Required]
        public byte[] PasswordHash { get; set; } = null!;

        [Required]
        public byte[] PasswordSalt { get; set; } = null!;

        public DateTime CreatedAt { get; set; } = DateTime.UtcNow;
    }

    public class PasswordResetToken
    {
        public int Id { get; set; }
        public int UserId { get; set; }
        public User? User { get; set; }
        
        [Required]
        public string Token { get; set; } = null!;
        
        public DateTime ExpiresAt { get; set; }
        public bool IsUsed { get; set; } = false;
        public DateTime CreatedAt { get; set; } = DateTime.UtcNow;
    }
}
