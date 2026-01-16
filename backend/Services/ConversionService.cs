using System;
using System.IO;
using System.IO.Compression;
using System.Threading.Tasks;
using Microsoft.Extensions.Logging;
using Spire.Pdf; // Keep for PdfToWord etc. if needed, or remove if unused (kept for now for existing PDF to X implementations)
using PDFtoImage;
using SkiaSharp;
using System.Diagnostics;

namespace ConverterApi.Services
{
    public interface IConversionService
    {
        Task<byte[]> ConvertAsync(Stream inputStream, string fileName, string sourceFormat, string targetFormat);
    }

    public class ConversionService : IConversionService
    {
        private readonly ILogger<ConversionService> _logger;

        public ConversionService(ILogger<ConversionService> logger)
        {
            _logger = logger;
        }

        public async Task<byte[]> ConvertAsync(Stream inputStream, string fileName, string sourceFormat, string targetFormat)
        {
            var source = sourceFormat.ToLower().TrimStart('.');
            var target = targetFormat.ToLower().TrimStart('.');

            _logger.LogInformation("Dönüştürme başlıyor: {FileName} ({Source} -> {Target})", fileName, source, target);

            return (source, target) switch
            {
                ("docx" or "doc", "pdf") => await WordToPdfAsync(inputStream),
                ("xlsx" or "xls", "pdf") => await ExcelToPdfAsync(inputStream),
                ("pptx" or "ppt", "pdf") => await PowerPointToPdfAsync(inputStream),
                ("pdf", "docx") => await PdfToWordAsync(inputStream),
                ("pdf", "xlsx") => await PdfToExcelAsync(inputStream),
                ("pdf", "pptx") => await PdfToPowerPointAsync(inputStream),
                ("pdf", "png") => await PdfToImageAsync(inputStream, SKEncodedImageFormat.Png),
                ("pdf", "jpg" or "jpeg") => await PdfToImageAsync(inputStream, SKEncodedImageFormat.Jpeg),
                _ => throw new NotSupportedException($"'{source}' formatından '{target}' formatına dönüştürme henüz desteklenmiyor.")
            };
        }

        private async Task<byte[]> WordToPdfAsync(Stream inputStream)
        {
            return await ConvertWithLibreOfficeAsync(inputStream, "docx");
        }

        private async Task<byte[]> ExcelToPdfAsync(Stream inputStream)
        {
             return await ConvertWithLibreOfficeAsync(inputStream, "xlsx");
        }

        private async Task<byte[]> PowerPointToPdfAsync(Stream inputStream)
        {
             return await ConvertWithLibreOfficeAsync(inputStream, "pptx");
        }

        private async Task<byte[]> ConvertWithLibreOfficeAsync(Stream inputStream, string inputExtension)
        {
            /*
             * LibreOffice Headless Conversion logic:
             * 1. Save input stream to temp file
             * 2. Run soffice --headless --convert-to pdf --outdir /tmp input.docx
             * 3. Read output pdf
             * 4. Cleanup
             */
            
            var tempId = Guid.NewGuid().ToString();
            var inputPath = Path.Combine(Path.GetTempPath(), $"{tempId}.{inputExtension}");
            var outputDir = Path.GetTempPath();
            var outputPath = Path.Combine(outputDir, $"{tempId}.pdf");

            try 
            {
                using (var fileStream = File.Create(inputPath))
                {
                    await inputStream.CopyToAsync(fileStream);
                }

                var processStartInfo = new ProcessStartInfo
                {
                    FileName = "soffice", // libreoffice or soffice depending on distro, usually soffice works if 'libreoffice' package installed
                    Arguments = $"--headless --convert-to pdf --outdir \"{outputDir}\" \"{inputPath}\"",
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    UseShellExecute = false,
                    CreateNoWindow = true
                };
                
                // Fallback for Mac (dev env) if 'soffice' not in PATH, mostly for production Linux
                // Only for PRODUCTION (Linux) this is guaranteed to work if Dockerfile updated.
                // On Local Mac without LibreOffice installed this will fail.
                
                _logger.LogInformation("Executing LibreOffice conversion: {FileName}", inputPath);

                using var process = new Process { StartInfo = processStartInfo };
                process.Start();
                var error = await process.StandardError.ReadToEndAsync();
                await process.WaitForExitAsync();

                if (process.ExitCode != 0)
                {
                    throw new Exception($"LibreOffice conversion failed. ExitCode: {process.ExitCode}. Error: {error}");
                }

                if (!File.Exists(outputPath))
                {
                     throw new FileNotFoundException($"LibreOffice did not generate output file at {outputPath}. Error: {error}");
                }

                return await File.ReadAllBytesAsync(outputPath);

            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "LibreOffice conversion error");
                throw;
            }
            finally
            {
                if (File.Exists(inputPath)) File.Delete(inputPath);
                if (File.Exists(outputPath)) File.Delete(outputPath);
            }
        }

        private async Task<byte[]> PdfToWordAsync(Stream inputStream)
        {
             return await Task.Run(() =>
            {
                using var pdf = new PdfDocument();
                pdf.LoadFromStream(inputStream);
                using var output = new MemoryStream();
                pdf.SaveToStream(output, Spire.Pdf.FileFormat.DOCX);
                return output.ToArray();
            });
        }

        private async Task<byte[]> PdfToExcelAsync(Stream inputStream)
        {
            return await Task.Run(() =>
            {
                using var pdf = new PdfDocument();
                pdf.LoadFromStream(inputStream);
                using var output = new MemoryStream();
                pdf.SaveToStream(output, Spire.Pdf.FileFormat.XLSX);
                return output.ToArray();
            });
        }

        private async Task<byte[]> PdfToPowerPointAsync(Stream inputStream)
        {
            return await Task.Run(() =>
            {
                using var pdf = new PdfDocument();
                pdf.LoadFromStream(inputStream);
                using var output = new MemoryStream();
                pdf.SaveToStream(output, Spire.Pdf.FileFormat.PPTX);
                return output.ToArray();
            });
        }

#pragma warning disable CA1416
        private async Task<byte[]> PdfToImageAsync(Stream inputStream, SKEncodedImageFormat format)
        {
            return await Task.Run(() =>
            {
                using var ms = new MemoryStream();
                inputStream.CopyTo(ms);
                var pdfBytes = ms.ToArray();

                var pageCount = Conversion.GetPageCount(pdfBytes);
                if (pageCount == 1)
                {
                    using var output = new MemoryStream();
                    if (format == SKEncodedImageFormat.Png)
                        Conversion.SavePng(output, pdfBytes);
                    else
                        Conversion.SaveJpeg(output, pdfBytes);
                    return output.ToArray();
                }
                else
                {
                    using var zipMs = new MemoryStream();
                    using (var archive = new ZipArchive(zipMs, ZipArchiveMode.Create, true))
                    {
                        for (int i = 0; i < pageCount; i++)
                        {
                            var ext = format == SKEncodedImageFormat.Png ? "png" : "jpg";
                            var entry = archive.CreateEntry($"sayfa_{i + 1}.{ext}");
                            using var entryStream = entry.Open();
                            if (format == SKEncodedImageFormat.Png)
                                Conversion.SavePng(entryStream, pdfBytes, page: i);
                            else
                                Conversion.SaveJpeg(entryStream, pdfBytes, page: i);
                        }
                    }
                    return zipMs.ToArray();
                }
            });
        }
#pragma warning restore CA1416
    }
}
