using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xunit;

namespace SpreadsheetLight.InsertRow
{
    public class InsertSpecRow
    {
        [Fact]
        public void ShouldInsertRow()
        {
            SLDocument sl = new SLDocument();

            for (int i = 1; i < 25; ++i)
            {
                for (int j = 1; j < 15; ++j)
                {
                    sl.SetCellValue(i, j, string.Format("R{0}C{1}", i, j));
                }
            }

            // insert 4 rows at row 3
            sl.InsertRow(3, 4);

            // delete 2 rows at row 10
            // Note that the original row 6 is now at row 10
            //sl.DeleteRow(10, 2);

            // insert 3 columns at column 5
            //sl.InsertColumn(5, 3);

            // delete 1 column at column 6
            // Note that columns 5, 6 and 7 are now blank because of the insert column operation
            // above. So this is deleting the middle column of the newly added 3 columns.
            //sl.DeleteColumn(6, 1);

            sl.SaveAs("InsertDeleteRowColumn.xlsx");
        }
    }
}
