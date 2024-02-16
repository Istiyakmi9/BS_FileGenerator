using iText.Html2pdf;

namespace BS_FileGenerator.Service
{
    public class HtmlToPdfConverter
    {

        public void ConvertToPdf(string html, string filePath)
        {
            try
            {
                FileStream pdfStream = new FileStream(filePath, FileMode.Create);
                HtmlConverter.ConvertToPdf(html, pdfStream);
            }
            catch (Exception)
            {
                throw;
            }
        }
    }
}
