using Spire.Xls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Xls.Core.Spreadsheet.PivotTables;

namespace ClearPivotFields
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

            // Get the sheet in which the pivot table is located
            Worksheet sheet = workbook.Worksheets["PivotTable"];

            // Get the first pivot table from the sheet
            XlsPivotTable pt = sheet.PivotTables[0] as XlsPivotTable;

            // Clear all the data fields in the pivot table
            pt.DataFields.Clear();

            // Calculate the pivot table data
            pt.CalculateData();

            // Specify the filename for the resulting workbook
            string result = "ClearPivotFields_result.xlsx";

            // Save the modified workbook to the file
            workbook.SaveToFile(result, ExcelVersion.Version2010);

            // Dispose of the workbook object to release resources
            workbook.Dispose();

            //View the document
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
