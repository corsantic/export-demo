using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using ClosedXML.Excel;
using ExcelExportDemo.model;
using Microsoft.Extensions.Options;

namespace ExcelExportDemo.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class HomeController : ControllerBase
    {
        private readonly AppSettings _appSettings;

        private List<User> users = new List<User>
        {
            new User {Id = 1, Username = "DoloresAbernathy"},
            new User {Id = 2, Username = "MaeveMillay"},
            new User {Id = 3, Username = "BernardLowe"},
            new User {Id = 4, Username = "ManInBlack"}
        };


        public HomeController(IOptions<AppSettings> appSettings)
        {
            _appSettings = appSettings.Value;
        }

        [HttpGet("xlsx")]
        public IActionResult Excel()
        {
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add(_appSettings.WorkSheetOptions.WorkSheetTitle);
                var currentRow = 1;
                var column = 1;
                var type = typeof(User);
                ConfigureFreezeColumns(worksheet);
                ConfigureFreezeRows(worksheet);


                var properties = SetHeaders(type, worksheet, currentRow, column);


                SetValues(currentRow, properties, worksheet);

                using (var stream = new MemoryStream())
                {
                    workbook.SaveAs(stream);
                    var content = stream.ToArray();

                    return File(
                        content,
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        "users.xlsx");
                }
            }
        }

        private void ConfigureFreezeRows(IXLWorksheet worksheet)
        {
            foreach (var freezeRow in _appSettings.WorkSheetOptions.FreezeRows)
                worksheet.SheetView.FreezeRows(freezeRow);
        }

        private void ConfigureFreezeColumns(IXLWorksheet worksheet)
        {
            foreach (var freezeColumn in _appSettings.WorkSheetOptions.FreezeColumns)
                worksheet.SheetView.FreezeColumns(freezeColumn);
        }

        private void SetValues(int currentRow, PropertyInfo[] properties, IXLWorksheet worksheet)
        {
            int column;
            foreach (var user in users)
            {
                currentRow++;
                column = 1;
                foreach (var property in properties)
                {
                    worksheet.Cell(currentRow, column).Value = property.GetValue(user, null);
                    column++;
                }
            }
        }

        private static PropertyInfo[] SetHeaders(Type type, IXLWorksheet worksheet, int currentRow, int column)
        {
            var properties = type.GetProperties();

            foreach (var prop in properties)
            {
                worksheet.Cell(currentRow, column).Value = prop.Name;
                column++;
            }

            return properties;
        }


        [HttpGet("csv")]
        public IActionResult Csv()
        {
            var builder = new StringBuilder();
            builder.AppendLine("Id,Username");
            foreach (var user in users) builder.AppendLine($"{user.Id},{user.Username}");

            return File(Encoding.UTF8.GetBytes(builder.ToString()), "text/csv", "users.csv");
        }
    }
}