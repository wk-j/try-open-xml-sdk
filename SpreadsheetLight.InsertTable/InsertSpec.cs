using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xunit;

namespace SpreadsheetLight.InsertTable
{
    public class InsertSpec
    {
        [Fact]
        public void ShouldInsertTable()
        {
            SLDocument sl = new SLDocument();

            int i, j;
            for (i = 2; i <= 12; ++i)
            {
                for (j = 2; j <= 6; ++j)
                {
                    if (i == 2)
                    {
                        sl.SetCellValue(i, j, string.Format("Col{0}", j));
                    }
                    else
                    {
                        sl.SetCellValue(i, j, i * j);
                    }
                }
            }

            // tabular data ranges from B2:F12, inclusive of a header row
            SLTable tbl = sl.CreateTable("B2", "F12");
            tbl.SetTableStyle(SLTableStyleTypeValues.Medium9);
            sl.InsertTable(tbl);

            for (i = 2; i <= 12; ++i)
            {
                for (j = 9; j <= 15; ++j)
                {
                    if (i == 2)
                    {
                        sl.SetCellValue(i, j, string.Format("Col{0}", j));
                    }
                    else
                    {
                        sl.SetCellValue(i, j, i * j);
                    }
                }
            }

            tbl = sl.CreateTable("I2", "O12");

            tbl.HasTotalRow = true;
            // 1st table column, column I
            tbl.SetTotalRowLabel(1, "Totals");
            // 7th table column, column O
            tbl.SetTotalRowFunction(7, SLTotalsRowFunctionValues.Sum);
            tbl.SetTableStyle(SLTableStyleTypeValues.Dark4);

            tbl.HasBandedColumns = true;
            tbl.HasBandedRows = true;
            tbl.HasFirstColumnStyled = true;
            tbl.HasLastColumnStyled = true;

            // sort by the 3rd table column (column K) in descending order
            tbl.Sort(3, false);

            sl.InsertTable(tbl);

            sl.SaveAs("Tables.xlsx");

        }
    }
}
