using System.Threading.Tasks;
using MailKit.Net.Smtp;
using MimeKit;
using Microsoft.Extensions.Configuration;

namespace ConverterApi.Services
{
    public interface IEmailService
    {
        Task SendPasswordResetEmailAsync(string toEmail, string resetToken, string resetLink);
        Task SendVerificationEmailAsync(string toEmail, string verificationLink);
    }

    public class EmailService : IEmailService
    {
        private readonly IConfiguration _config;
        private readonly Microsoft.Extensions.Logging.ILogger<EmailService> _logger;

        public EmailService(IConfiguration config, Microsoft.Extensions.Logging.ILogger<EmailService> logger)
        {
            _config = config;
            _logger = logger;
        }

        public async Task SendPasswordResetEmailAsync(string toEmail, string resetToken, string resetLink)
        {
            var smtpServer = _config["Email:SmtpServer"]?.Trim();
            var smtpPort = int.Parse(_config["Email:SmtpPort"]?.Trim() ?? "587");
            var fromEmail = _config["Email:FromEmail"]?.Trim();
            var fromName = _config["Email:FromName"]?.Trim();
            var smtpUsername = _config["Email:SmtpUsername"]?.Trim();
            var smtpPassword = _config["Email:SmtpPassword"]?.Trim();

            var message = new MimeMessage();
            message.From.Add(new MailboxAddress(fromName, fromEmail));
            message.To.Add(new MailboxAddress(toEmail, toEmail));
            message.Subject = "Şifre Sıfırlama Talebi";

            var bodyBuilder = new BodyBuilder
            {
                HtmlBody = $@"
                    <h2>Şifre Sıfırlama</h2>
                    <p>Merhaba,</p>
                    <p>Şifre sıfırlama talebiniz alınmıştır. Aşağıdaki kodu kullanarak şifrenizi sıfırlayabilirsiniz:</p>
                    <p><strong>Sıfırlama Kodu: {resetToken}</strong></p>
                    <p>Bu kod 1 saat geçerlidir.</p>
                    <p>Saygılarımızla,<br/>Converter Ekibi</p>"
            };

            message.Body = bodyBuilder.ToMessageBody();

            try
            {
                using (var client = new SmtpClient())
                {
                    client.Timeout = 5000;
                    var options = smtpPort == 465 ? MailKit.Security.SecureSocketOptions.SslOnConnect : MailKit.Security.SecureSocketOptions.StartTls;
                    await client.ConnectAsync(smtpServer, smtpPort, options);
                    await client.AuthenticateAsync(smtpUsername, smtpPassword);
                    await client.SendAsync(message);
                    await client.DisconnectAsync(true);
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "STMP Error: Password reset failed for {Email}", toEmail);
            }
        }

        public async Task SendVerificationEmailAsync(string toEmail, string verificationLink)
        {
            var smtpServer = _config["Email:SmtpServer"]?.Trim();
            var smtpPort = int.Parse(_config["Email:SmtpPort"]?.Trim() ?? "587");
            var fromEmail = _config["Email:FromEmail"]?.Trim();
            var fromName = _config["Email:FromName"]?.Trim();
            var smtpUsername = _config["Email:SmtpUsername"]?.Trim();
            var smtpPassword = _config["Email:SmtpPassword"]?.Trim();

            var message = new MimeMessage();
            message.From.Add(new MailboxAddress(fromName, fromEmail));
            message.To.Add(new MailboxAddress(toEmail, toEmail));
            message.Subject = "Email Doğrulama";

            var bodyBuilder = new BodyBuilder
            {
                HtmlBody = $@"
                    <h2>Email Doğrulama</h2>
                    <p>Merhaba,</p>
                    <p>Kaydınızı tamamlamak için lütfen aşağıdaki linke tıklayın:</p>
                    <p><a href='{verificationLink}'>Hesabımı Doğrula</a></p>
                    <p>veya bu linki tarayıcınıza yapıştırın: {verificationLink}</p>
                    <p>Saygılarımızla,<br/>Converter Ekibi</p>"
            };

            message.Body = bodyBuilder.ToMessageBody();

            try
            {
                using (var client = new SmtpClient())
                {
                    client.Timeout = 5000;
                    var options = smtpPort == 465 ? MailKit.Security.SecureSocketOptions.SslOnConnect : MailKit.Security.SecureSocketOptions.StartTls;
                    await client.ConnectAsync(smtpServer, smtpPort, options);
                    await client.AuthenticateAsync(smtpUsername, smtpPassword);
                    await client.SendAsync(message);
                    await client.DisconnectAsync(true);
                }
                _logger.LogInformation("Verification Email sent successfully to {Email}", toEmail);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "STMP Error: Verification Email failed for {Email}. Message: {Message}", toEmail, ex.Message);
            }
        }
    }
}
