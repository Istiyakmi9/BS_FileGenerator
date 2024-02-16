using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using HtmlToOpenXml;

namespace BS_FileGenerator.Service
{
    public class HtmlToDocxConverter
    {
        public string ToDocx(string html, string destinationFolder, string headerLogoPath)
        {
            string empty = string.Empty;
            if (!string.IsNullOrEmpty(html))
            {
                if (File.Exists(destinationFolder))
                {
                    File.Delete(destinationFolder);
                }

                using MemoryStream memoryStream = new MemoryStream();
                using (WordprocessingDocument wordprocessingDocument = WordprocessingDocument.Create((Stream)memoryStream, WordprocessingDocumentType.Document))
                {
                    MainDocumentPart mainDocumentPart = wordprocessingDocument.MainDocumentPart;
                    if (mainDocumentPart == null)
                    {
                        mainDocumentPart = wordprocessingDocument.AddMainDocumentPart();
                        new Document(new Body()).Save(mainDocumentPart);
                    }

                    HtmlConverter htmlConverter = new HtmlConverter(mainDocumentPart);
                    //if (File.Exists(headerLogoPath))
                    //{
                    //    htmlConverter.ProcessHeaderImage(mainDocumentPart, headerLogoPath);
                    //}

                    htmlConverter.ParseHtml(html);
                    mainDocumentPart.Document.Save();
                }

                File.WriteAllBytes(destinationFolder, memoryStream.ToArray());
                return empty;
            }

            return "Template path not found";
        }
    }
}
