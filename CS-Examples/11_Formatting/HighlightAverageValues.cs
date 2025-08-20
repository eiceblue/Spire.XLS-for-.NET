using System;
using System.Data.OleDb;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using Spire.Xls;
using Spire.Xls.Core.Spreadsheet;
using Spire.Xls.Core.Spreadsheet.Collections;
using Spire.Xls.Core;

namespace HighlightAverageValues
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

            // Load the file from disk
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\Template_Xls_6.xlsx");

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add conditional format to highlight cells below average values
            XlsConditionalFormats format1 = sheet.ConditionalFormats.Add();
            // Set the cell range to apply the formatting
            format1.AddRange(sheet.Range["E2:E10"]);
            // Add below average condition
            IConditionalFormat cf1 = format1.AddAverageCondition(AverageType.Below);
            // Set background color for cells below average
            cf1.BackColor = Color.SkyBlue; 

            // Add conditional format to highlight cells above average values
            XlsConditionalFormats format2 = sheet.ConditionalFormats.Add();
            // Set the cell range to apply the formatting
            format2.AddRange(sheet.Range["E2:E10"]);
            // Add above average condition
            IConditionalFormat cf2 = format2.AddAverageCondition(AverageType.Above);
            // Set background color for cells above average
            cf2.BackColor = Color.Orange;

            // Save the workbook to file
            String result = "Result-HighlightBelowAndAboveAverageValues.xlsx";
            workbook.SaveToFile(result, ExcelVersion.Version2013);

            // Dispose of the workbook object to release resources
            workbook.Dispose();

            //Launch the MS Excel file.
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
