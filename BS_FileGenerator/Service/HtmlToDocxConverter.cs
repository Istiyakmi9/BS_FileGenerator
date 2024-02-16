using BS_FileGenerator.IService;
using BS_FileGenerator.Models;
using DocMaker.HtmlToDocx;

namespace BS_FileGenerator.Service
{
    public class HtmlToDocxConverter : IHtmlToDocxConverter
    {
        private readonly IHTMLConverter _hTMLConverter;

        public HtmlToDocxConverter(IHTMLConverter hTMLConverter)
        {
            _hTMLConverter = hTMLConverter;
        }

        public async Task<byte[]> ToDocxAsync(RequestModel requestModel)
        {
            string empty = string.Empty;
            MemoryStream memoryStream = new MemoryStream();

            if (!string.IsNullOrEmpty(requestModel.Content))
            {
                memoryStream = await _hTMLConverter.ToDocxStreamAsync(requestModel.Content);
            }

            return await Task.FromResult(memoryStream.ToArray());
        }
    }
}