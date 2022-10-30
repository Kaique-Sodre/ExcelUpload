using Microsoft.AspNetCore.Http;
using OfficeOpenXml;

namespace ExcelUpload.Interface
{
    public interface IExcelService
    {
        ExcelWorksheet ReadXlsx(IFormFile file);
    }
}
