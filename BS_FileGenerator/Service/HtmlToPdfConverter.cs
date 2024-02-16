using BS_FileGenerator.IService;
using BS_FileGenerator.Models;
using DinkToPdf;
using DinkToPdf.Contracts;

namespace BS_FileGenerator.Service
{
    public class HtmlToPdfConverter : IHtmlToPdfConverter
    {
        private readonly IConverter _converter;

        public HtmlToPdfConverter(IConverter converter)
        {
            _converter = converter;
        }

        public async Task<MemoryStream> ConvertToPdfAsync(RequestModel requestModel)
        {
            var file = await GeneratePdfFile(requestModel);
            var memoryStream = new MemoryStream(file);
            return memoryStream;
        }

        public async Task<byte[]> GenerateHtmlToPdfBytesAsync(RequestModel requestModel)
        {
            return await GeneratePdfFile(requestModel);
        }

        private async Task<byte[]> GeneratePdfFile(RequestModel requestModel)
        {
            var globalSettings = new GlobalSettings
            {
                PaperSize = PaperKind.A4,
                Orientation = DinkToPdf.Orientation.Portrait,
                Margins = new MarginSettings { Top = 10, Bottom = 10, Left = 10, Right = 10 },
                DocumentTitle = "Generated PDF"
            };

            var objectSettings = new ObjectSettings
            {
                PagesCount = true,
                HtmlContent = requestModel.Content,
                WebSettings = { DefaultEncoding = "utf-8", UserStyleSheet = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "styles", "styles.css") },
                HeaderSettings = { FontSize = 9, Right = "Page [page] of [toPage]", Line = true },
                FooterSettings = { FontSize = 9, Line = true, Center = "Footer" }
            };

            var pdf = new HtmlToPdfDocument()
            {
                GlobalSettings = globalSettings,
                Objects = { objectSettings }
            };

            var file = _converter.Convert(pdf);
            return await Task.FromResult(file);
        }
    }
}
