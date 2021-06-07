using System;
using System.Data.OleDb;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using Spire.Xls;
using Spire.Xls.Core.Spreadsheet;

namespace ApplySubscriptAndSuperscript
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

            //Get the first worksheet.
			Worksheet sheet = workbook.Worksheets[0];

            sheet.Range["B2"].Text = "This is an example of Subscript:";
            sheet.Range["D2"].Text = "This is an example of Superscript:";

            //Set the rtf value of "B3" to "R100-0.06".
            CellRange range = sheet.Range["B3"];
            range.RichText.Text = "R100-0.06";

            //Create a font. Set the IsSubscript property of the font to "true".
            ExcelFont font = workbook.CreateFont();
            font.IsSubscript = true;
            font.Color = Color.Green;

            //Set font for specified range of the text in "B3".
            range.RichText.SetFont(4, 8, font);

            //Set the rtf value of "D3" to "a2 + b2 = c2".
            range = sheet.Range["D3"];
            range.RichText.Text = "a2 + b2 = c2";

            //Create a font. Set the IsSuperscript property of the font to "true".
            font = workbook.CreateFont();
            font.IsSuperscript = true;

            //Set font for specified range of the text in "D3".
            range.RichText.SetFont(1, 1, font);
            range.RichText.SetFont(6, 6, font);
            range.RichText.SetFont(11, 11, font);

            sheet.AllocatedRange.AutoFitColumns();

            String result = "Result-ApplySubscriptAndSuperscript.xlsx";

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
