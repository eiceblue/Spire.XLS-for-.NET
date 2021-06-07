using System;
using System.Data.OleDb;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using Spire.Xls;
using Spire.Xls.Core.Spreadsheet;
using Spire.Xls.Core.Spreadsheet.Shapes;

namespace RemoveBorderlineOfTextbox
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
            workbook.Version = ExcelVersion.Version2013;

            //Create a new worksheet named "Remove Borderline" and add a chart to the worksheet.
			Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Remove Borderline";
            Chart chart = sheet.Charts.Add();

            //Create textbox1 in the chart and input text information.
            XlsTextBoxShape textbox1 = chart.TextBoxes.AddTextBox(50, 50, 100, 600) as XlsTextBoxShape;
            textbox1.Text = "The solution with borderline";

            //Create textbox2 in the chart, input text information and remove borderline.
            XlsTextBoxShape textbox2 = chart.TextBoxes.AddTextBox(1000, 50, 100, 600) as XlsTextBoxShape;
            textbox2.Text = "The solution without borderline";
            textbox2.Line.Weight = 0; 

            String result = "Result-RemoveBorderlineOfTextbox.xlsx";

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
