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

namespace HighlightDuplicateUniqueValues
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

            //Use conditional formatting to highlight duplicate values in range "C2:C10" with IndianRed color.
            XlsConditionalFormats xcfs = sheet.ConditionalFormats.Add();
            xcfs.AddRange(sheet.Range["C2:C10"]);
            IConditionalFormat format1 = xcfs.AddCondition();
            format1.FormatType = ConditionalFormatType.DuplicateValues;
            format1.BackColor = Color.IndianRed;

            //Use conditional formatting to highlight unique values in range "C2:C10" with Yellow color.
            XlsConditionalFormats xcfs1 = sheet.ConditionalFormats.Add();
            xcfs1.AddRange(sheet.Range["C2:C10"]);
            IConditionalFormat format2 = xcfs.AddCondition(); 
            format2.FormatType = ConditionalFormatType.UniqueValues;
            format2.BackColor = Color.Yellow;

            String result = "Result-HighlightDuplicateAndUniqueValues.xlsx";

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
