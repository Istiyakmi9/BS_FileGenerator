using BS_FileGenerator.Models;

namespace BS_FileGenerator.IService
{
    public interface IHtmlToDocxConverter
    {
        Task<byte[]> ToDocxAsync(RequestModel requestModel);
    }
}