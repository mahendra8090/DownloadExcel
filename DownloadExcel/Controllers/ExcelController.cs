using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading.Tasks;

namespace DownloadExcel.Controllers
{
    [ApiController]
    [Route("[controller]")]

    public class ExcelController : Controller
    {


        [HttpPost]
        public async Task<IActionResult> ExportToExcel(int id)
        {
            try
            {

                List<Student> list = new List<Student>();
                for (int i = 0; i < 1000000; i++)
                {
                    list.Add(new Student()
                    {
                        StudentId = i,
                        FirstName = i.ToString(),
                    });
                }
                int sheets = list.Count % 100000==0? list.Count / 100000 : (list.Count / 100000) +1;
                int skip = 0;
                // Create a new Excel package using EPPlus
                using (var package = new ExcelPackage())
                {
                    for(int i = 0; i < sheets; i++)
                    {
                        var records = list.Skip(skip * 100000).Take(100000);
                        var worksheet = package.Workbook.Worksheets.Add($"Sheet_{i}");

                        // Add headers to the Excel file
                        worksheet.Cells[1, 1].Value = "ID";
                        worksheet.Cells[1, 2].Value = "Name";
                        // Add more headers as needed

                        // Add data to the Excel file
                        int rowIndex = 2;
                        foreach (var record in records)
                        {
                            worksheet.Cells[rowIndex, 1].Value =record.StudentId; // Replace with your field names
                            worksheet.Cells[rowIndex, 2].Value = record.FirstName;
                            // Add more fields as needed

                            rowIndex++;
                        }
                        skip++;
                    }

                    Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
                    // Generate the Excel file as a byte array
                    var excelData = package.GetAsByteArray();

                    // Stream the Excel file to the client using FileStreamResult
                    var stream = new MemoryStream(excelData);
                    return new FileStreamResult(stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                    {
                        FileDownloadName = "data.xlsx"
                    };
                }
            }
            catch (Exception ex)
            {
                // Handle any exceptions here, e.g., log the error or return an error response
                return BadRequest("An error occurred while exporting data to Excel.");
            }
        }

        public class Student
        {
            public int StudentId { get; set; }
            public string FirstName { get; set; }
            public string LastName { get; set; }
            public DateTime DateOfBirth { get; set; }
            public string Email { get; set; }
            public string PhoneNumber { get; set; }
        }
    }


}