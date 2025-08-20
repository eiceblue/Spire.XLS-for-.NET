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

namespace CreateFormulaConditionalFormat
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

            //Get the first worksheet and the first column from the workbook.
			Worksheet sheet = workbook.Worksheets[0];
            CellRange range = sheet.Columns[0];

            //Set the conditional formatting formula and apply the rule to the chosen cell range.
            XlsConditionalFormats xcfs = sheet.ConditionalFormats.Add();
            xcfs.AddRange(range);
            IConditionalFormat conditional = xcfs.AddCondition();
            conditional.FormatType = ConditionalFormatType.Formula;
            conditional.FirstFormula = "=($A1<$B1)";
            conditional.BackKnownColor = ExcelColors.Yellow;

            String result = "Result-CreateFormulaToApplyConditionalFormatting.xlsx";

            //Save to file.
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
