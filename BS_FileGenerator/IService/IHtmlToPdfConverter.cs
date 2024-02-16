using BS_FileGenerator.Models;

namespace BS_FileGenerator.IService
{
    public interface IHtmlToPdfConverter
    {
        Task<MemoryStream> ConvertToPdfAsync(RequestModel requestModel);
        Task<byte[]> GenerateHtmlToPdfBytesAsync(RequestModel requestModel);
    }
}
