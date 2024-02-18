using Bot.CoreBottomHalf.CommonModal.API;
using BS_FileGenerator.Controller;
using BS_FileGenerator.IService;
using BS_FileGenerator.Models;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using ModalLayer.Modal;

namespace BS_FileGenerator.Controllers
{
    [Route("api/generate/[controller]")]
    [ApiController]
    public class FileGeneratorController : BaseController
    {
        private readonly IHtmlToDocxConverter _htmlToDocxConverter;
        private readonly IHtmlToPdfConverter _htmlToPdfConverter;
        private readonly IDataTableToExcel _dataTableToExcel;
        private readonly ILogger<FileGeneratorController> _logger;

        public FileGeneratorController(IDataTableToExcel dataTableToExcel,
            IHtmlToPdfConverter htmlToPdfConverter,
            IHtmlToDocxConverter htmlToDocxConverter,
            ILogger<FileGeneratorController> logger)
        {
            _dataTableToExcel = dataTableToExcel;
            _htmlToPdfConverter = htmlToPdfConverter;
            _htmlToDocxConverter = htmlToDocxConverter;
            _logger = logger;
        }

        #region PDF Handler


        [HttpPost("generate_pdf")]
        [AllowAnonymous]
        public async Task<IActionResult> HtmlToPdfConverter(RequestModel requestModel)
        {
            try
            {
                var file = await _htmlToPdfConverter.ConvertToPdfAsync(requestModel);
                return new FileStreamResult(file, "application/pdf")
                {
                    FileDownloadName = "generated.pdf"
                };
            }
            catch (Exception ex)
            {
                _logger.LogError(ex.Message);
                throw;
            }
        }

        [HttpPost("generate_pdf_bytes")]
        [AllowAnonymous]
        public async Task<byte[]> GenerateHtmlToPdfBytes(RequestModel requestModel)
        {
            try
            {
                return await _htmlToPdfConverter.GenerateHtmlToPdfBytesAsync(requestModel);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex.Message);
                throw;
            }
        }


        #endregion

        #region DOCX Handler

        [HttpPost("generate_docx")]
        [AllowAnonymous]
        public async Task<IActionResult> HtmlToDocConverter(RequestModel requestModel)
        {
            try
            {
                var file = await _htmlToDocxConverter.ToDocxAsync(requestModel);
                return File(file, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "generated.docx");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex.Message);
                throw;
            }
        }

        [HttpPost("generate_docx_bytes")]
        [AllowAnonymous]
        public async Task<byte[]> GenerateHtmlToDocBytes(RequestModel requestModel)
        {
            try
            {
                return await _htmlToDocxConverter.ToDocxAsync(requestModel);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex.Message);
                throw;
            }
        }

        #endregion


        #region EXCEL Handler

        [HttpPost("generate_excel")]
        public async Task DatatableToExcelConverter()
        {
            try
            {
                _dataTableToExcel.ToExcel(null, "");
                await Task.CompletedTask;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex.Message);
                throw;
            }
        }


        [HttpPost("generate_excel_bytes")]
        public async Task GenerateDatatableToExcelBytes()
        {
            try
            {
                _dataTableToExcel.ToExcel(null, "");
                await Task.CompletedTask;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex.Message);
                throw;
            }
        }


        #endregion
    }
}
