using Spire.Xls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace ReplaceAndHighlight
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Load the workbook from file
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\ReplaceAndHighlight.xlsx");

            // Get the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Find all occurrences of the string "Total" in the worksheet, including case-sensitive and whole word matches
            CellRange[] ranges = worksheet.FindAllString("Total", true, true);

            // Iterate through each found range
            foreach (CellRange range in ranges)
            {
                // Reset the text in the range by replacing it with "Sum"
                range.Text = "Sum";

                // Set the color of the range to yellow
                range.Style.Color = Color.Yellow;
            }

            // Specify the file name for the resulting workbook after replacement and highlighting
            string result = "ReplaceAndHighlight_result.xlsx";

            // Save the modified workbook to the specified file in Excel 2010 format
            workbook.SaveToFile(result, ExcelVersion.Version2010);

            // Dispose of the workbook object to release resources
            workbook.Dispose();

            // Lauch the result file
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
