using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Newtonsoft.Json;

namespace ConsoleApp1
{
    public static class ExportExcel
    {
        public static Stream CreateExcel(IEnumerable<IExcelReport> list, string sheetName = "گزارش")
        {
            var excelFileName = Path.GetTempPath() + DateTime.Now.Ticks + ".xlsx";

            var table = (DataTable) JsonConvert.DeserializeObject(JsonConvert.SerializeObject(list),
                (typeof(DataTable)));

            var headerDictionary = TypeDescriptor.GetProperties(list.GetType().GetGenericArguments().Single())
                .Cast<PropertyDescriptor>()
                .ToDictionary(p => p.Name, p => p.DisplayName);

            using (var document = SpreadsheetDocument.Create(excelFileName, SpreadsheetDocumentType.Workbook))
            {
                var workbookPart = document.AddWorkbookPart();
                workbookPart.Workbook = new Workbook();

                var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                var sheetData = new SheetData();
                worksheetPart.Worksheet = new Worksheet(sheetData);

                var sheets = workbookPart.Workbook.AppendChild(new Sheets());
                var sheet = new Sheet() {Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = sheetName};

                sheets.Append(sheet);

                var headerRow = new Row();

                var columns = new List<string>();
                foreach (DataColumn column in table.Columns)
                {
                    columns.Add(column.ColumnName);
                    headerRow.AppendChild(new Cell
                    {
                        DataType = CellValues.String,
                        CellValue = new CellValue(headerDictionary[column.ColumnName])
                    });
                }

                sheetData.AppendChild(headerRow);

                foreach (DataRow drowse in table.Rows)
                {
                    var newRow = new Row();
                    foreach (var cell in columns.Select(col =>
                                 new Cell
                                 {
                                     DataType = CellValues.String, CellValue = new CellValue(drowse[col].ToString())
                                 }))
                        newRow.AppendChild(cell);

                    sheetData.AppendChild(newRow);
                }

                workbookPart.Workbook.Save();

            }

            var fileByte = File.ReadAllBytes(excelFileName);
            File.Delete(excelFileName);
            return new MemoryStream(fileByte);
        }
    }
}
