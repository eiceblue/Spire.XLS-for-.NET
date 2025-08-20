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

            // Apply conditional formatting to highlight duplicate values in the range "C2:C10" with the color IndianRed.
            XlsConditionalFormats duplicateFormats = sheet.ConditionalFormats.Add();
            duplicateFormats.AddRange(sheet.Range["C2:C10"]);
            IConditionalFormat duplicateCondition = duplicateFormats.AddCondition();
            duplicateCondition.FormatType = ConditionalFormatType.DuplicateValues;
            duplicateCondition.BackColor = Color.IndianRed;

            // Apply conditional formatting to highlight unique values in the range "C2:C10" with the color Yellow.
            XlsConditionalFormats uniqueFormats = sheet.ConditionalFormats.Add();
            uniqueFormats.AddRange(sheet.Range["C2:C10"]);
            IConditionalFormat uniqueCondition = uniqueFormats.AddCondition();
            uniqueCondition.FormatType = ConditionalFormatType.UniqueValues;
            uniqueCondition.BackColor = Color.Yellow;

            // Specify the output file name.
            String result = "Result-HighlightDuplicateAndUniqueValues.xlsx";

            // Save the workbook to a file using Excel 2013 format.
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
