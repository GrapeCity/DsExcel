using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples.Features.Tables
{
    public class InsertDeleteTableRowColumns : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Excel.Workbook workbook)
        {
            object[,] data = new object[,]{
                {"Name", "City", "Birthday", "Sex", "Weight", "Height"},
                {"Bob", "NewYork", new DateTime(1968, 6, 8), "male", 80, 180},
                {"Betty", "NewYork", new DateTime(1972, 7, 3), "female", 72, 168},
                {"Gary", "NewYork", new DateTime(1964, 3, 2), "male", 71, 179},
                {"Hunk", "Washington", new DateTime(1972, 8, 8), "male", 80, 171},
                {"Cherry", "Washington", new DateTime(1986, 2, 2), "female", 58, 161},
                { "Eva", "Washington", new DateTime(1993, 2, 15), "female", 71, 180}
            };

            IWorksheet worksheet = workbook.Worksheets[0];
            worksheet.Range["A1:F7"].Value = data;
            worksheet.Range["A:F"].ColumnWidth = 15;

            //add table.
            ITable table = worksheet.Tables.Add(worksheet.Range["A1:F7"], true);

            //add table column before first column.
            table.Columns.Add(0);
            //add table column before second column.
            table.Columns.Add(1);

            //delete first table column.
            table.Columns[0].Delete();
            //delete "City" table column.
            table.Columns["City"].Delete();

            //insert a table row in table's last row.
            table.Rows.Add();
            //delete second table row.
            table.Rows[1].Delete();            
        }
    }
}
