using ExcelUpload.Interface;
using Microsoft.AspNetCore.Http;
using OfficeOpenXml;
using System.IO;

namespace ExcelToDatabase.Services
{
    public class ExcelService : IExcelService
    {
        public ExcelWorksheet ReadXlsx(IFormFile file)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using MemoryStream stream = new();
            file.CopyTo(stream);

            ExcelPackage package = new(stream);
            ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
            
            return worksheet;
        }
    }
}
