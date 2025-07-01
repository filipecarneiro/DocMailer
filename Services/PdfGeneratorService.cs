using DocMailer.Models;
using Markdig;
using iText.Html2pdf;
using iText.Html2pdf.Resolver.Font;
using iText.Kernel.Pdf;
using iText.Layout;
using System.Globalization;
using System.Text;

namespace DocMailer.Services
{
    /// <summary>
    /// Service for generating PDF documents
    /// </summary>
    public class PdfGeneratorService
    {
        private readonly MarkdownPipeline _markdownPipeline;

        public PdfGeneratorService()
        {
            _markdownPipeline = new MarkdownPipelineBuilder()
                .UseAdvancedExtensions()
                .Build();
        }

        public string GeneratePdf(string markdownContent, string outputPath, Recipient recipient, DocumentTemplate template)
        {
            // Convert Markdown to HTML
            var html = Markdown.ToHtml(markdownContent, _markdownPipeline);
            
            // Generate header HTML if logo is specified
            var headerHtml = GenerateHeaderHtml(template.Metadata);
            
            // Load CSS content from file or use default
            var cssContent = LoadCssContent(template.Metadata);
            
            // Generate complete HTML
            var styledHtml = $@"
<!DOCTYPE html>
<html>
<head>
    <meta charset='UTF-8'>
    <style>
{cssContent}
    </style>
</head>
<body>
    {headerHtml}
    <div class='content'>
        {html}
    </div>
</body>
</html>";

            // Generate filename based on template title and recipient name
            var documentTitle = template.Metadata.ContainsKey("title") ? 
                template.Metadata["title"].ToString() ?? "Document" : "Document";
            var safeTitle = SanitizeFileName(documentTitle).Replace(" ", "_");
            var safeName = SanitizeFileName(recipient.Name).Replace(" ", "_");
            var fileName = $"{safeTitle}-{safeName}.pdf";
            var fullPath = Path.Combine(outputPath, fileName);

            // Get page size from metadata (default to A4)
            var pageSize = GetPageSize(template.Metadata);

            // Convert HTML to PDF with custom properties
            using var pdfWriter = new PdfWriter(fullPath);
            
            // Create converter properties to set page size
            var converterProperties = new ConverterProperties();
            
            // Convert HTML to PDF
            using var pdfDoc = new PdfDocument(pdfWriter);
            
            // Set custom PDF properties
            var info = pdfDoc.GetDocumentInfo();
            info.SetTitle(documentTitle);
            info.SetAuthor(template.Metadata.ContainsKey("author") ? 
                template.Metadata["author"].ToString() ?? "DocMailer" : "DocMailer");
            info.SetSubject(template.Metadata.ContainsKey("type") ? 
                template.Metadata["type"].ToString() ?? "" : "");
            info.SetProducer("DocMailer by Filipe Carneiro");
            
            // Set page size
            pdfDoc.SetDefaultPageSize(pageSize);
            
            // Convert HTML to PDF with document
            var document = HtmlConverter.ConvertToDocument(styledHtml, pdfDoc, converterProperties);
            document.Close();

            return fullPath;
        }

        private string SanitizeFileName(string fileName)
        {
            if (string.IsNullOrEmpty(fileName))
                return "Document";

            // Normalize to remove accents/diacritics from any language
            var normalized = fileName.Normalize(NormalizationForm.FormD);
            var stringBuilder = new StringBuilder();

            foreach (var c in normalized)
            {
                var unicodeCategory = CharUnicodeInfo.GetUnicodeCategory(c);
                if (unicodeCategory != UnicodeCategory.NonSpacingMark)
                {
                    stringBuilder.Append(c);
                }
            }

            var withoutAccents = stringBuilder.ToString().Normalize(NormalizationForm.FormC);

            // Remove invalid file name characters
            var invalidChars = Path.GetInvalidFileNameChars();
            var sanitized = new string(withoutAccents.Where(c => !invalidChars.Contains(c)).ToArray());
            
            return sanitized.Trim();
        }

        private iText.Kernel.Geom.PageSize GetPageSize(Dictionary<string, object> metadata)
        {
            if (metadata.ContainsKey("pageSize"))
            {
                var pageSize = metadata["pageSize"].ToString()?.ToUpperInvariant();
                return pageSize switch
                {
                    "A4" => iText.Kernel.Geom.PageSize.A4,
                    "A3" => iText.Kernel.Geom.PageSize.A3,
                    "A5" => iText.Kernel.Geom.PageSize.A5,
                    "LETTER" => iText.Kernel.Geom.PageSize.LETTER,
                    "LEGAL" => iText.Kernel.Geom.PageSize.LEGAL,
                    _ => iText.Kernel.Geom.PageSize.A4
                };
            }
            return iText.Kernel.Geom.PageSize.A4;
        }

        private string GenerateHeaderHtml(Dictionary<string, object> metadata)
        {
            if (!metadata.ContainsKey("headerCenter"))
                return "";

            var logoPath = metadata["headerCenter"].ToString();
            if (string.IsNullOrEmpty(logoPath))
                return "";

            // Check if logo file exists
            if (!File.Exists(logoPath))
            {
                // Try relative to current directory
                var relativePath = Path.Combine(Directory.GetCurrentDirectory(), logoPath);
                if (!File.Exists(relativePath))
                    return "";
                logoPath = relativePath;
            }

            // Convert image to base64 for embedding
            try
            {
                var imageBytes = File.ReadAllBytes(logoPath);
                var base64Image = Convert.ToBase64String(imageBytes);
                var mimeType = GetImageMimeType(logoPath);
                
                return $@"
                <div class='header'>
                    <img src='data:{mimeType};base64,{base64Image}' alt='Logo' />
                </div>";
            }
            catch
            {
                // If image processing fails, return empty header
                return "";
            }
        }

        private string GetImageMimeType(string filePath)
        {
            var extension = Path.GetExtension(filePath).ToLowerInvariant();
            return extension switch
            {
                ".png" => "image/png",
                ".jpg" or ".jpeg" => "image/jpeg",
                ".gif" => "image/gif",
                ".svg" => "image/svg+xml",
                ".bmp" => "image/bmp",
                _ => "image/png"
            };
        }

        private string LoadCssContent(Dictionary<string, object> metadata)
        {
            // Check if custom stylesheet is specified
            if (metadata.ContainsKey("styleSheet"))
            {
                var cssPath = metadata["styleSheet"].ToString();
                if (!string.IsNullOrEmpty(cssPath))
                {
                    // Try to load the specified CSS file
                    if (File.Exists(cssPath))
                    {
                        try
                        {
                            return File.ReadAllText(cssPath);
                        }
                        catch
                        {
                            // If loading fails, fall back to default
                        }
                    }
                    
                    // Try relative to current directory
                    var relativePath = Path.Combine(Directory.GetCurrentDirectory(), cssPath);
                    if (File.Exists(relativePath))
                    {
                        try
                        {
                            return File.ReadAllText(relativePath);
                        }
                        catch
                        {
                            // If loading fails, fall back to default
                        }
                    }
                }
            }

            // Return default CSS if no custom stylesheet or loading failed
            return GetDefaultCss();
        }

        private string GetDefaultCss()
        {
            return @"
        body { 
            font-family: Arial, sans-serif; 
            margin: 40px;
            line-height: 1.6;
        }
        h1, h2, h3 { 
            color: #333; 
            text-align: center;
        }
        .header {
            text-align: center;
            padding-bottom: 30px;
            margin-bottom: 30px;
        }
        .header img {
            max-width: 200px;
            max-height: 80px;
            height: auto;
        }
        .content {
            margin-top: 20px;
        }
        .footer {
            border-top: 1px solid #ccc;
            padding-top: 20px;
            margin-top: 30px;
            font-size: 12px;
            color: #666;
        }";
        }
    }
}
