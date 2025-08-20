using Spire.Xls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Xls.Core.Spreadsheet.PivotTables;

namespace UpdateDataSource
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

            // Access the "Data" worksheet
            Worksheet data = workbook.Worksheets["Data"];

            // Modify the data source by changing the value in cell A2 to "NewValue"
            data.Range["A2"].Text = "NewValue";

            // Modify the data source by changing the value in cell D2 to 28000
            data.Range["D2"].NumberValue = 28000;

            // Access the worksheet containing the pivot table
            Worksheet sheet = workbook.Worksheets["PivotTable"];

            // Get the first pivot table from the worksheet
            XlsPivotTable pt = sheet.PivotTables[0] as XlsPivotTable;

            // Set the pivot table's cache to refresh on load
            pt.Cache.IsRefreshOnLoad = true;

            // Calculate and update the pivot table data
            pt.CalculateData();

            // Specify the filename for the updated workbook
            String result = "UpdateDataSource_result.xlsx";

            // Save the modified workbook to a file (in Excel 2010 format)
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
