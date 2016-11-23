using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Xunit;


namespace DocumentFormat
{
    public class ReplaceTextSpec
    {
        [Fact]
        public void ShouldReplaceTextInWord()
        {
            var year = 2015;
            var month = 10;
            var day = 25;

            var document = @"Input\ReplaceText.docx";

            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(document, true))
            {
                string docText = null;
                using (StreamReader sr = new StreamReader(wordDoc.MainDocumentPart.GetStream()))
                {
                    docText = sr.ReadToEnd();
                }

                docText = docText
                    .Replace("{year}", $"{year}")
                    .Replace("{month}", $"{month}")
                    .Replace("{day}", $"{day}");

                using (StreamWriter sw = new StreamWriter(wordDoc.MainDocumentPart.GetStream(FileMode.Create)))
                {
                    sw.Write(docText);
                }
            }
        }

        [Fact]
        public void ShouldReplaceTextInExcel()
        {
            var fileName = @"Input\ReplaceText.xlsx";

            using (var fs = new FileStream(fileName, FileMode.Open, FileAccess.ReadWrite, FileShare.ReadWrite))
            {
                using (var doc = SpreadsheetDocument.Open(fs, true))
                {
                    var workbookPart = doc.WorkbookPart;
                    var sstPart = workbookPart.GetPartsOfType<SharedStringTablePart>().First();
                    //var sst = sstpart.SharedStringTable;

                    // Iterate through all the items in the SharedStringTable. If the text already exists, return its index.
                    foreach (var item in sstPart.SharedStringTable.Elements<SharedStringItem>())
                    {
                        if (item.InnerText.ToString().Contains("{year}"))
                        {
                            Text child = item.Descendants<Text>().Where(x => x.InnerText == "{year}").FirstOrDefault();
                            if (child != null)
                            {
                                child.Text = "2016";
                            }
                        }
                    }
                    sstPart.SharedStringTable.Save();
                }
            }
        }
    }
}
