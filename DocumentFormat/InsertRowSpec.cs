using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xunit;

namespace DocumentFormat
{
    public class InsertRowSpec
    {
        [Fact]
        public void ShouldInsertRow()
        {
            var fileName = @"Input\InsertRow.xlsx";

            var items = new List<Dictionary<string, string>>();
            items.Add(new Dictionary<string, string> { { "aa", "bb" } });

            using (var fs = new FileStream(fileName, FileMode.Open, FileAccess.ReadWrite, FileShare.ReadWrite))
            {
                using (var doc = SpreadsheetDocument.Open(fs, true))
                {
                    var workbookPart = doc.WorkbookPart;
                    //SheetData sheetData = workbookPart.Workbook.GetFirstChild<SheetData>();
                    var sheetData = workbookPart.WorksheetParts.First().Worksheet.GetFirstChild<SheetData>();
                    var first = sheetData.Elements<Row>().First();

                    if (first != null)
                    {
                        sheetData.InsertBefore(new Row() { RowIndex = (first.RowIndex + 1) }, first);
                    }
                    else
                    {
                        var row = sheetData.InsertAt(new Row() { Height = 300 } , 0);
                    }

                    workbookPart.Workbook.Save();
                }
            }
        }

        // Got in valid output
        [Fact]
        public void ShouldCloneRow()
        {
            var fileName = @"Input\CloneRow.xlsx";

            var items = new List<Dictionary<string, string>>();
            items.Add(new Dictionary<string, string> { { "aa", "bb" } });

            using (var fs = new FileStream(fileName, FileMode.Open, FileAccess.ReadWrite, FileShare.ReadWrite))
            {
                using (var doc = SpreadsheetDocument.Open(fs, true))
                {
                    var workbookPart = doc.WorkbookPart;
                    var sheetData = workbookPart.WorksheetParts.First().Worksheet.GetFirstChild<SheetData>();
                    var first = sheetData.Elements<Row>().First();

                    var last = sheetData.Elements<Row>().Last();

                    if (first != null)
                    {
                        var newRow = first.CloneNode(true) as Row;
                        //sheetData.InsertAt(newRow, (int) last.RowIndex.Value);
                        sheetData.Append(newRow);
                    }
                    else
                    {
                        var row = sheetData.InsertAt(new Row() { Height = 300 }, 0);
                    }

                    workbookPart.Workbook.Save();
                }
            }
        }
    }
}
