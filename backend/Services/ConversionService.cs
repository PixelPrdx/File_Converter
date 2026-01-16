using System;
using System.IO;
using System.IO.Compression;
using System.Threading.Tasks;
using Microsoft.Extensions.Logging;
using Spire.Doc;
using Spire.Xls;
using Spire.Presentation;
using Spire.Pdf;
using PDFtoImage;
using SkiaSharp;

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
            return await Task.Run(() =>
            {
                try
                {
                    using var doc = new Document();
                    doc.LoadFromStream(inputStream, Spire.Doc.FileFormat.Auto);
                    using var output = new MemoryStream();
                    doc.SaveToStream(output, Spire.Doc.FileFormat.PDF);
                    return output.ToArray();
                }
                catch (Exception ex)
                {
                    _logger.LogError(ex, "Word to PDF conversion failed.");
                    throw;
                }
            });
        }

        private async Task<byte[]> ExcelToPdfAsync(Stream inputStream)
        {
            return await Task.Run(() =>
            {
                using var workbook = new Workbook();
                workbook.LoadFromStream(inputStream);
                using var output = new MemoryStream();
                workbook.SaveToStream(output, Spire.Xls.FileFormat.PDF);
                return output.ToArray();
            });
        }

        private async Task<byte[]> PowerPointToPdfAsync(Stream inputStream)
        {
            return await Task.Run(() =>
            {
                using var presentation = new Presentation();
                presentation.LoadFromStream(inputStream, Spire.Presentation.FileFormat.Auto);
                using var output = new MemoryStream();
                presentation.SaveToFile(output, Spire.Presentation.FileFormat.PDF);
                return output.ToArray();
            });
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
