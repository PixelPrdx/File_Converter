using System.Threading.Tasks;
using MailKit.Net.Smtp;
using MimeKit;
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

        public EmailService(IConfiguration config, ILogger<EmailService> logger)
        {
            _config = config;
            _logger = logger;
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
                await SendEmailInternalAsync(toEmail, "SMTP Test Mesajı", "Bu bir test mesajıdır. Bu mesajı görüyorsanız SMTP ayarlarınız doğrudur.");
                return true;
            }
            catch (System.Exception ex)
            {
                _logger.LogError(ex, "STMP Test failed for {Email}", toEmail);
                throw;
            }
        }

        private async Task SendEmailInternalAsync(string toEmail, string subject, string htmlBody)
        {
            var smtpServer = _config["Email:SmtpServer"]?.Trim();
            var smtpPortStr = _config["Email:SmtpPort"]?.Trim();
            var smtpPort = string.IsNullOrEmpty(smtpPortStr) ? 587 : int.Parse(smtpPortStr);
            var fromEmail = _config["Email:FromEmail"]?.Trim();
            var fromName = _config["Email:FromName"]?.Trim();
            var smtpUsername = _config["Email:SmtpUsername"]?.Trim();
            var smtpPassword = _config["Email:SmtpPassword"]?.Trim();

            var message = new MimeMessage();
            message.From.Add(new MailboxAddress(fromName, fromEmail));
            message.To.Add(new MailboxAddress(toEmail, toEmail));
            message.Subject = subject;

            var bodyBuilder = new BodyBuilder { HtmlBody = htmlBody };
            message.Body = bodyBuilder.ToMessageBody();

            using (var client = new SmtpClient())
            {
                // Render connectivity optimizations
                client.Timeout = 20000; // Increased to 20s
                client.CheckCertificateRevocation = false;

                // Port based security selection
                var options = MailKit.Security.SecureSocketOptions.Auto;
                if (smtpPort == 465) options = MailKit.Security.SecureSocketOptions.SslOnConnect;
                else if (smtpPort == 587) options = MailKit.Security.SecureSocketOptions.StartTls;

                _logger.LogInformation("Connecting to SMTP {Host}:{Port} ({Options})", smtpServer, smtpPort, options);

                try
                {
                    await client.ConnectAsync(smtpServer, smtpPort, options);
                    await client.AuthenticateAsync(smtpUsername, smtpPassword);
                    await client.SendAsync(message);
                    await client.DisconnectAsync(true);
                    _logger.LogInformation("Email '{Subject}' sent successfully to {Email}", subject, toEmail);
                }
                catch (System.Exception ex)
                {
                    _logger.LogError(ex, "Failed to send email '{Subject}' to {Email}. Port: {Port}", subject, toEmail, smtpPort);
                    throw;
                }
            }
        }
    }
}
