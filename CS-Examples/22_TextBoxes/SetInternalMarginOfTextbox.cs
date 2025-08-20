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

namespace SetInternalMarginOfTextbox
{
	public partial class Form1 : Form
	{
        public Form1()
        {
            InitializeComponent();
        }

		private void btnRun_Click(object sender, System.EventArgs e)
		{
            // Create a new workbook object
            Workbook workbook = new Workbook();

            // Load the Excel document from disk
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\Template_Xls_4.xlsx");

            // Get the first worksheet from the workbook
            Worksheet sheet = workbook.Worksheets[0];

            // Add a textbox to the sheet and set its position and size
            XlsTextBoxShape textbox = sheet.TextBoxes.AddTextBox(4, 2, 100, 300) as XlsTextBoxShape;

            // Set the text on the textbox
            textbox.Text = "Insert TextBox in Excel and set the margin for the text";
            textbox.HAlignment = CommentHAlignType.Center;
            textbox.VAlignment = CommentVAlignType.Center;

            // Set the inner margins of the contents
            textbox.InnerLeftMargin = 1;
            textbox.InnerRightMargin = 3;
            textbox.InnerTopMargin = 1;
            textbox.InnerBottomMargin = 1;

            // Specify the output filename for the workbook
            String result = "Result-SetInternalMarginOfTextbox.xlsx";

            // Save the modified workbook to a file
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
