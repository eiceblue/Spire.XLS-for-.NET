using Spire.Xls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Xls.Core.Spreadsheet;

namespace GetConditionalFormatColor
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            // Create a new workbook object
            Workbook workbook = new Workbook();

            // Load an existing Excel document
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\Template_Xls_13.xlsx");

            // Get the first sheet from the workbook
            Worksheet sheet = workbook.Worksheets[0];

            // Define a cell range
            CellRange cRange = sheet.Range["A1:C1"];

            // Retrieve the color of the condition format applied to the cell range
            var color = cRange.GetConditionFormatsStyle().Color;

            // Display a message box with the color information
            MessageBox.Show("The color of the condition format is " + color.ToString());

            // Dispose of the workbook object to release resources
            workbook.Dispose();
        }


        private void btnClose_Click(object sender, EventArgs e)
        {
            Close();
        }
    }
}
