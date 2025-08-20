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
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet from the workbook
            Worksheet sheet = workbook.Worksheets[0];

            // Set the text for cell B2
            sheet.Range["B2"].Text = "This is an example of Subscript:";
            // Set the text for cell D2
            sheet.Range["D2"].Text = "This is an example of Superscript:"; 

            // Set the RTF value of cell "B3" to "R100-0.06"
            CellRange range = sheet.Range["B3"];
            range.RichText.Text = "R100-0.06";

            // Create a font and set the IsSubscript property to true
            ExcelFont font = workbook.CreateFont();
            font.IsSubscript = true;
            font.Color = Color.Green;

            // Set the font for the specified range of text in cell "B3"
            range.RichText.SetFont(4, 8, font);

            // Set the RichText value of cell "D3" to "a2 + b2 = c2"
            range = sheet.Range["D3"];
            range.RichText.Text = "a2 + b2 = c2";

            // Create a font and set the IsSuperscript property to true
            font = workbook.CreateFont();
            font.IsSuperscript = true;

            // Set the font for the specified range of text in cell "D3"
            range.RichText.SetFont(1, 1, font);
            range.RichText.SetFont(6, 6, font);
            range.RichText.SetFont(11, 11, font);

            // Auto-fit the columns to adjust their widths
            sheet.AllocatedRange.AutoFitColumns();

            // Specify the name for the resulting Excel file
            String result = "Result-ApplySubscriptAndSuperscript.xlsx"; 

            // Save the modified workbook to a file
            workbook.SaveToFile(result, ExcelVersion.Version2013);

            // Dispose of the workbook object
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
