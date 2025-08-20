using Spire.Xls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Xls.Core.Spreadsheet.PivotTables;

namespace SetFormatOptions
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            //Create a workbook
            Workbook workbook = new Workbook();

            //Load an excel file including pivot table
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\PivotTableExample.xlsx");

            // Get the sheet where the pivot table is located
            Worksheet sheet = workbook.Worksheets["PivotTable"];

            // Access the first pivot table in the sheet
            XlsPivotTable pt = sheet.PivotTables[0] as XlsPivotTable;

            // Enable automatic formatting for the pivot table report
            pt.Options.IsAutoFormat = true;

            // Show grand totals for rows in the pivot table report
            pt.ShowRowGrand = true;

            // Show grand totals for columns in the pivot table report
            pt.ShowColumnGrand = true;

            // Display a custom string in cells that contain null values
            pt.DisplayNullString = true;
            pt.NullString = "null";

            // Set the layout of the pivot table report
            pt.PageFieldOrder = PagesOrderType.DownThenOver;

            // Specify the filename for the resulting workbook
            string result = "SetFormatOptions_result.xlsx";

            // Save the modified workbook to a file
            workbook.SaveToFile(result, ExcelVersion.Version2010);

            // Dispose of the workbook object to release resources
            workbook.Dispose();

            // Launch the document
            FileViewer(result);
        }

        private void FileViewer(string fileName)
        {
            try
            {
                System.Diagnostics.Process.Start(fileName);
            }
            catch { }
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            Close();
        }
    }
}
