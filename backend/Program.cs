using System.Text;
using ConverterApi.Data;
using ConverterApi.Services;
using Microsoft.AspNetCore.Authentication.JwtBearer;
using Microsoft.EntityFrameworkCore;
using Microsoft.IdentityModel.Tokens;

var builder = WebApplication.CreateBuilder(args);

// Ensure Environment Variables are explicitly added
builder.Configuration.AddEnvironmentVariables();

var frontendUrl = (builder.Configuration["FrontendUrl"] 
                 ?? builder.Configuration["FRONTEND_URL"] 
                 ?? "https://file-converter-phi-nine.vercel.app").Trim();

if (frontendUrl.EndsWith("/")) frontendUrl = frontendUrl.TrimEnd('/');

builder.Services.AddCors(options =>
{
    options.AddPolicy("FrontendOnly", policy =>
    {
        policy.WithOrigins(frontendUrl, "http://localhost:8000")
              .AllowAnyMethod()
              .AllowAnyHeader();
    });
});

builder.Services.AddScoped<IEmailService, EmailService>();
builder.Services.AddControllers();
builder.Services.AddEndpointsApiExplorer();
builder.Services.AddSwaggerGen();

var cs = builder.Configuration.GetConnectionString("DefaultConnection");
Console.WriteLine("CONNSTR LEN = " + (cs?.Length ?? 0));
builder.Services.AddDbContext<AppDbContext>(opt =>
    opt.UseSqlServer(cs, sql =>
        sql.EnableRetryOnFailure(
            maxRetryCount: 10,
            maxRetryDelay: TimeSpan.FromSeconds(30),
            errorNumbersToAdd: null)));




var key = Encoding.UTF8.GetBytes(builder.Configuration["Jwt:Key"]!);
builder.Services.AddAuthentication(options =>
{
    options.DefaultAuthenticateScheme = JwtBearerDefaults.AuthenticationScheme;
    options.DefaultChallengeScheme = JwtBearerDefaults.AuthenticationScheme;
}).AddJwtBearer(options =>
{
    options.TokenValidationParameters = new TokenValidationParameters
    {
        ValidateIssuer = true,
        ValidateAudience = true,
        ValidateIssuerSigningKey = true,
        ValidIssuer = builder.Configuration["Jwt:Issuer"],
        ValidAudience = builder.Configuration["Jwt:Audience"],
        IssuerSigningKey = new SymmetricSecurityKey(key)
    };
});

var app = builder.Build();

// Migration'ı arka planda çalıştırarak startup timeout'u engelle (Hata olsa bile uygulama açık kalsın)
_ = Task.Run(async () => {
    using (var scope = app.Services.CreateScope())
    {
        var services = scope.ServiceProvider;
        var logger = services.GetRequiredService<ILogger<Program>>();
        var db = services.GetRequiredService<AppDbContext>();
        try 
        { 
            var connectionString = app.Configuration.GetConnectionString("DefaultConnection");
            logger.LogInformation("Background migration starting...");
            
            if (!string.IsNullOrEmpty(connectionString)) {
                var builder = new Microsoft.Data.SqlClient.SqlConnectionStringBuilder(connectionString);
                logger.LogInformation("Attempting connection to Host: {DataSource}, DB: {InitialCatalog}", builder.DataSource, builder.InitialCatalog);
            }

            // Render'ın dış dünyaya hangi IP ile çıktığını öğrenmek için:
            try {
                using var client = new HttpClient();
                var outboundIp = await client.GetStringAsync("https://api.ipify.org");
                logger.LogInformation("Render Server Outbound IP (for Azure Firewall): {OutboundIp}", outboundIp.Trim());
            } catch (Exception ex) {
                logger.LogWarning("Could not determine outbound IP: {Message}", ex.Message);
            }

            await db.Database.MigrateAsync(); 
            logger.LogInformation("Database migration completed successfully.");
        } 
        catch (Exception ex) 
        { 
            logger.LogError(ex, "Background migration failed. Confirm environment variables are set correctly in Render.");
        }
    }
});

// Swagger'ı prod'da da görmek istersen bunu kaldırma:
app.UseSwagger();
app.UseSwaggerUI();

app.UseRouting();
app.UseCors("FrontendOnly");
app.UseAuthentication();
app.UseAuthorization();
app.MapControllers();

// Render'ın gerçek IP'sini öğrenmek için X-Forwarded-For başlığına bakan kritik endpoint:
app.MapGet("/ip", (HttpContext ctx) => {
    var forwarded = ctx.Request.Headers["X-Forwarded-For"].FirstOrDefault();
    if (!string.IsNullOrEmpty(forwarded)) return Results.Ok(new { ip = forwarded.Split(',')[0] });
    return Results.Ok(new { ip = ctx.Connection.RemoteIpAddress?.ToString() });
});

app.MapGet("/config-debug", (IConfiguration config) => 
{
    var keys = config.AsEnumerable().Select(k => k.Key).ToList();
    return Results.Ok(new { 
        FoundKeys = keys,
        FrontendUrl = config["FrontendUrl"] ?? config["FRONTEND_URL"] ?? "Still Null",
        SmtpServer = config["Email:SmtpServer"] ?? "Still Null"
    });
});

app.Run();
