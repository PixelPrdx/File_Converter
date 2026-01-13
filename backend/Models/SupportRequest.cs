using System;

namespace ConverterApi.Models
{
    public class SupportRequest
    {
        public int Id { get; set; }
        public int? UserId { get; set; }
        public string Subject { get; set; } = string.Empty;
        public string Message { get; set; } = string.Empty;
        public DateTime CreatedAt { get; set; } = DateTime.UtcNow;
    }
}
