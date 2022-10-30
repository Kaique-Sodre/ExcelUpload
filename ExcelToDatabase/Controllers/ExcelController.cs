using ExcelToDatabase.Models;
using ExcelUpload.Interface;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
using System;
using System.Collections.Generic;

namespace ExcelToDatabase.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class ExcelController : ControllerBase
    {
        private readonly IExcelService _excelService;   
        public ExcelController(IExcelService excelService){
            _excelService = excelService;
        }

        [Consumes("multipart/form-data")]
        [HttpPost("upload-excel")]
        public ActionResult UploadExcel(IFormFile file)
        {
            //transform IFormFile to worksheet (CREATE SERVICE HERE)
            ExcelWorksheet worksheet = _excelService.ReadXlsx(file);

            IList<Product> data = new List<Product>();

            int rowCount = worksheet.Dimension.End.Row;

            //goes through each of sheet rows, create a product and push it on list
            for (int row = 2; row < rowCount; row++)
            {
                Product product = new()
                {
                    Id = Guid.NewGuid().ToString("N"),
                    Name = worksheet.Cells[row, 1].Value.ToString(),
                    Quantity = Convert.ToInt32(worksheet.Cells[row, 2].Value),
                    Value = Convert.ToDecimal(worksheet.Cells[row, 3].Value),
                    CreationDate = DateTime.Now
                };

                data.Add(product);
            }

            //uses the list to AddRange on database
            //call your product service here passing the list

            //return results
            return Ok("Upload successful.");
        }
    }
}