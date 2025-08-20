using System;
using System.Data.OleDb;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;

using Spire.Xls;
using Spire.Xls.Core.Spreadsheet.PivotTables;

namespace CreateTable
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnRun_Click(object sender, System.EventArgs e)
        {
            // Create a new workbook object
            Workbook workbook = new Workbook();  

            // Load the workbook from the specified file path
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\CreateTable.xlsx");

            // Get the first worksheet of the workbook
            Worksheet sheet = workbook.Worksheets[0];  

            // Add a new List Object to the worksheet with the name "table" and range [1, 1, 19, 5]
            sheet.ListObjects.Create("table", sheet.Range[1, 1, 19, 5]);

            // Apply a default style (TableStyleLight9) to the created table
            sheet.ListObjects[0].BuiltInTableStyle = TableBuiltInStyles.TableStyleLight9;

            // Specify the output file name
            string result = "CreateTable_out.xlsx"; 

            // Save the modified workbook to the specified file path in Excel 2010 format
            workbook.SaveToFile(result, ExcelVersion.Version2010);

            // Dispose of the workbook object to free up resources
            workbook.Dispose();

            // View file
            ExcelDocViewer(result);
        }

        private void ExcelDocViewer(string fileName)
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
