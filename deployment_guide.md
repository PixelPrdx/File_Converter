# Nova PDF Yayınlama Rehberi (Deployment Guide)

Nova PDF uygulamasını yerel bilgisayarından (localhost) internete taşımak için aşağıdaki adımları izleyebilirsin.

## 1. Hazırlık ve Yapılandırma

Uygulamayı her iki tarafın (Frontend ve Backend) birbirini bulabileceği şekilde yapılandırmalısın.

### Frontend Yapılandırması (`frontend/config.js`)
Yayınladığın yerdeki Backend URL'sini buraya yazmalısın:
```javascript
const CONFIG = {
    API_BASE: "https://api.senindomaine.com/api" // Üretim URL'si
};
```

### Backend Yapılandırması (`backend/appsettings.json`)
Buradaki `FrontendUrl` kısmını, web sitenin yayında olacağı adresle güncelle:
```json
{
  "FrontendUrl": "https://www.senindomainin.com",
  "ConnectionStrings": {
    "DefaultConnection": "Server=...;Database=..." // Prod DB bilgileri
  }
}
```

---

## 2. Yayınlama Seçenekleri

### Seçenek A: VPS (DigitalOcean, Vultr, Hetzner) - *Tavsiye Edilen*
Kendi sunucuna kurmak tam kontrol sağlar.

1.  **Sunucuya Bağlan**: SSH ile sunucuna gir.
2.  **Docker Kullan (Opsiyonel)**: Uygulamayı Dockerize ederek kolayca taşıyabilirsin.
3.  **Manuel Kurulum**: 
    - Sunucuya `.NET Runtime` kur.
    - `dotnet publish -c Release` komutuyla dosyaları hazırla ve sunucuya kopyala.
    - Nginx kullanarak Frontend dosyalarını servis et ve Backend'e yönlendirme (Proxy) yap.
    - **SSL (HTTPS)**: `Certbot` ile ücretsiz SSL kur.

### Seçenek B: Bulut Platformları (Ücretsiz/Düşük Maliyetli)
- **Frontend**: [Vercel](https://vercel.com) veya [Netlify](https://netlify.com) üzerine ücretsiz yükleyebilirsin. (Dosyaları sürükle-bırak yapman yeterli).
- **Backend**: [Render](https://render.com) veya [Railway](https://railway.app) ücretsiz .NET desteği sunar.

---

## 3. Önemli Güvenlik Notları
- **CORS**: `Program.cs` içindeki `AllowAll` politikasını, sadece kendi domainine izin verecek şekilde (`AllowAnyOrigin` yerine `.WithOrigins("https://domain.com")`) kısıtlamalısın.
- **Şifreler**: `appsettings.json` içindeki SMTP ve JWT anahtarlarını prodüksiyon için çok daha güçlü şifrelerle değiştirmelisin.
- **Database**: Canlı uygulamada SQLite yerine `PostgreSQL` veya `SQL Server` kullanmanı öneririm.

*Uygulaman artık internete çıkmaya hazır! Herhangi bir adımda yardıma ihtiyacın olursa buradayım.*
