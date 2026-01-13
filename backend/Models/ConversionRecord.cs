using System;
using System.ComponentModel.DataAnnotations;

namespace ConverterApi.Models
{
    public class ConversionRecord
    {
        public int Id { get; set; }

        public int UserId { get; set; }
        public User? User { get; set; }

        [Required]
        [MaxLength(255)]
        public string OriginalFileName { get; set; } = null!;

        [Required]
        [MaxLength(255)]
        public string ConvertedFileName { get; set; } = null!;

        [Required]
        public string FilePath { get; set; } = null!;

        public long FileSize { get; set; }

        public DateTime CreatedAt { get; set; } = DateTime.UtcNow;
    }
}
