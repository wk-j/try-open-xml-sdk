using Syncfusion.XlsIO;
using System.Collections.Generic;
using System.IO;
using Xunit;

namespace XlsIO
{
    public class Customer
    {
        public string SalesPerson { set; get; }
        public string SalesJanJune { set;get;}
        public byte[] Image { set; get; }
        public string SalesJulyDec { set; get; } = "Go go go";
        public string Change { set; get; }
        public string NumbersTable { set; get; }
    }

    public class NumberTable
    {
        public int Column0 { set; get; } = 10;
        public int Column1 { set; get; } = 20;
        public int Column2 { set; get; } = 30;
        public int Column3 { set; get; } = 40;
    }

    public class RenderSpec
    {
        [Fact]
        public void ShouldRenderXls()
        {
            var excelEngine = new ExcelEngine();
            excelEngine.Excel.DefaultVersion = ExcelVersion.Excel2007;

            var template = @"Input\TemplateMarkerImageWithSize&Position.xlsx";
            var imagePath = @"Input\Octocat.png";

            var workbook = excelEngine.Excel.Workbooks.Open(template);
            workbook.Version = ExcelVersion.Excel2007;

            var marker = workbook.CreateTemplateMarkersProcessor();
            var image = File.ReadAllBytes(imagePath);

            var customers = new List<Customer> {
                new Customer { SalesJanJune = "AAA", SalesPerson = "BBB", Image = image },
                new Customer { SalesJanJune = "BBB", SalesPerson = "BBB", Image = image },
                new Customer { SalesJanJune = "CCC", SalesPerson = "BBB", Image = image },
                new Customer { SalesJanJune = "DDD", SalesPerson = "BBB", Image = image },
            };

            var numbers = new List<NumberTable>
            {
                new NumberTable { },
                new NumberTable { },
                new NumberTable { }
            };

            marker.AddVariable("Customers", customers);
            marker.AddVariable("NumbersTable", numbers);

            marker.ApplyMarkers();

            workbook.SaveAs("TemplateMarkerImageWithSize&Position-Output.xlsx");
            workbook.Clone();
        }
    }
}
