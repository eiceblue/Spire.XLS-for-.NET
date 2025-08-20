using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;

using Spire.Xls;
using Spire.Xls.Core;

namespace AddScrollBarControl
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        private void btnRun_Click(object sender, System.EventArgs e)
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Load the Excel document from the specified file path
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\ExcelSample_N1.xlsx");

            // Get the first worksheet from the workbook
            Worksheet sheet = workbook.Worksheets[0];

            // Set a value and formatting for range B10
            sheet.Range["B10"].Value2 = 1;
            sheet.Range["B10"].Style.Font.IsBold = true;

            // Add a scroll bar control to the worksheet
            IScrollBarShape scrollBar = sheet.ScrollBarShapes.AddScrollBar(10, 3, 150, 20);
            // Link the scroll bar control to cell B10
            scrollBar.LinkedCell = sheet.Range["B10"];
            // Set the minimum value of the scroll bar
            scrollBar.Min = 1;
            // Set the maximum value of the scroll bar
            scrollBar.Max = 150;
            // Set the incremental change when scrolling
            scrollBar.IncrementalChange = 1;
            // Enable 3D shading for the scroll bar
            scrollBar.Display3DShading = true;

            //Specify the filename for the resulting Excel file
            string output = "AddScrollBarControl_out.xlsx";

            // Save the workbook to the specified file in Excel 2013 format
            workbook.SaveToFile(output, ExcelVersion.Version2013);

            // Dispose of the workbook object to release resources
            workbook.Dispose();

            //Launch the Excel file
            ExcelDocViewer(output);
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
