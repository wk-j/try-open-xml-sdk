using ReadWriteExcel.Lib;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xunit;

namespace ExcelSample
{
    public class ExcelSampleSpec
    {
        [Fact]
        public void ShouldWriteFile()
        {
            var writer = new SLExcelWriter();
            var result = writer.GenerateExcel(new SLExcelData
            {
                Headers = new List<string> { "F1", "F2", "F3" },
                DataRows = new List<List<string>>
                {
                    new List<string> { "F1 1", "F2 1", "F3 1" },
                    new List<string> { "F1 2", "F2 2", "F3 2" },
                    new List<string> { "F1 3", "F2 3", "F3 3" },
                }
            });
            File.WriteAllBytes("Excel.Output.xlsx", result);
        }
    }
}
