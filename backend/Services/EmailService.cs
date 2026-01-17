using System.Threading.Tasks;
using System.Net.Http;
using System.Net.Http.Json;
using System.Text.Json;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;

namespace ConverterApi.Services
{
    public interface IEmailService
    {
        Task SendPasswordResetEmailAsync(string toEmail, string resetToken, string resetLink);
        Task SendVerificationEmailAsync(string toEmail, string verificationLink);
        Task<bool> TestEmailAsync(string toEmail);
    }

    public class EmailService : IEmailService
    {
        private readonly IConfiguration _config;
        private readonly ILogger<EmailService> _logger;
        private readonly IHttpClientFactory _httpClientFactory;

        public EmailService(IConfiguration config, ILogger<EmailService> logger, IHttpClientFactory httpClientFactory)
        {
            _config = config;
            _logger = logger;
            _httpClientFactory = httpClientFactory;
        }

        public async Task SendPasswordResetEmailAsync(string toEmail, string resetToken, string resetLink)
        {
            var body = $@"
                <h2>Şifre Sıfırlama</h2>
                <p>Merhaba,</p>
                <p>Şifre sıfırlama talebiniz alınmıştır. Aşağıdaki kodu kullanarak şifrenizi sıfırlayabilirsiniz:</p>
                <p><strong>Sıfırlama Kodu: {resetToken}</strong></p>
                <p>Bu kod 1 saat geçerlidir.</p>
                <p>Saygılarımızla,<br/>Converter Ekibi</p>";

            await SendEmailInternalAsync(toEmail, "Şifre Sıfırlama Talebi", body);
        }

        public async Task SendVerificationEmailAsync(string toEmail, string verificationLink)
        {
            var body = $@"
                <h2>Email Doğrulama</h2>
                <p>Merhaba,</p>
                <p>Kaydınızı tamamlamak için lütfen aşağıdaki linke tıklayın:</p>
                <p><a href='{verificationLink}'>Hesabımı Doğrula</a></p>
                <p>veya bu linki tarayıcınıza yapıştırın: {verificationLink}</p>
                <p>Saygılarımızla,<br/>Converter Ekibi</p>";

            await SendEmailInternalAsync(toEmail, "Email Doğrulama", body);
        }

        public async Task<bool> TestEmailAsync(string toEmail)
        {
            try
            {
                await SendEmailInternalAsync(toEmail, "Brevo API Test Mesajı", "Bu bir test mesajıdır. Bu mesajı görüyorsanız Brevo HTTP API entegrasyonu başarılıdır.");
                return true;
            }
            catch (System.Exception ex)
            {
                _logger.LogError(ex, "Brevo API Test failed for {Email}", toEmail);
                throw;
            }
        }

        private async Task SendEmailInternalAsync(string toEmail, string subject, string htmlBody)
        {
            // API Settings from configuration
            // Priority: appsettings.json or Environment Variables (Brevo__ApiKey)
            var apiKey = _config["Brevo:ApiKey"] ?? _config["BREVO_API_KEY"];
            var senderEmail = _config["Brevo:SenderEmail"] ?? _config["BREVO_SENDER_EMAIL"] ?? "pixelprdx@gmail.com";
            var senderName = _config["Brevo:SenderName"] ?? "Nova PDF Converter";

            if (string.IsNullOrEmpty(apiKey))
            {
                _logger.LogWarning("Brevo API Key is missing. Falling back to SMTP config if available...");
                // Note: We could fallback to SMTP here, but since the goal is to skip SMTP, we throw if API key is missing.
                throw new System.Exception("Brevo API Key is not configured. Please set 'Brevo:ApiKey' in Render.");
            }

            var client = _httpClientFactory.CreateClient();
            var request = new HttpRequestMessage(HttpMethod.Post, "https://api.brevo.com/v3/smtp/email");
            request.Headers.Add("api-key", apiKey);

            var payload = new
            {
                sender = new { name = senderName, email = senderEmail },
                to = new[] { new { email = toEmail } },
                subject = subject,
                htmlContent = htmlBody
            };

            request.Content = JsonContent.Create(payload);

            _logger.LogInformation("Sending email via Brevo API to {Email}", toEmail);

            var response = await client.SendAsync(request);
            
            if (!response.IsSuccessStatusCode)
            {
                var error = await response.Content.ReadAsStringAsync();
                _logger.LogError("Brevo API Error: {StatusCode} - {Error}", response.StatusCode, error);
                throw new System.Exception($"Brevo API call failed: {response.StatusCode}. Details: {error}");
            }

            _logger.LogInformation("Email '{Subject}' sent successfully via Brevo API", subject);
        }
    }
}
