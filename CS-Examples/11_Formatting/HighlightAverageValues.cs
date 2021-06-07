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
            //Create a workbook.
			Workbook workbook = new Workbook();

            //Load the file from disk.
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\Template_Xls_6.xlsx");

            //Get the first worksheet.
			Worksheet sheet = workbook.Worksheets[0];

            //Add conditional format.
            XlsConditionalFormats format1 = sheet.ConditionalFormats.Add();
            //Set the cell range to apply the formatting.
            format1.AddRange(sheet.Range["E2:E10"]);
            //Add below average condition.
            IConditionalFormat cf1 = format1.AddAverageCondition(AverageType.Below);
            //Highlight cells below average values.
            cf1.BackColor = Color.SkyBlue;

            //Add conditional format.
            XlsConditionalFormats format2 = sheet.ConditionalFormats.Add();
            //Set the cell range to apply the formatting.
            format2.AddRange(sheet.Range["E2:E10"]);
            //Add above average condition.
            IConditionalFormat cf2 = format1.AddAverageCondition(AverageType.Above);
            //Highlight cells above average values.
            cf2.BackColor = Color.Orange;

            String result = "Result-HighlightBelowAndAboveAverageValues.xlsx";

            //Save to file.
            workbook.SaveToFile(result, ExcelVersion.Version2013);

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
