using Spire.Xls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Xls.Core.Spreadsheet;

namespace TextDirection
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            // Create a workbook
            Workbook workbook = new Workbook();
       
            // Add a new worksheet to the Excel object
            Worksheet sheet = workbook.Worksheets[0];

            // Access the "B5" cell from the worksheet
            CellRange cell = sheet.Range["B5"];

            // Add some value to the "B5" cell
            cell.Text = "Hello Spire!";

            // Set the reading order from right to left of the text in the "B5" cell
            cell.Style.ReadingOrder = ReadingOrderType.RightToLeft;

            // Specify the name for the resulting Excel file
            String result = "TextDirection_result.xlsx";

            // Save the workbook to the specified file in Excel 2010 format
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
