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
using System.IO;
using System.Text;

namespace ExtractTextFromATextbox
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
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\Template_Xls_5.xlsx");

            //Get the first worksheet.
			Worksheet sheet = workbook.Worksheets[0];

            //Get the first textbox.
            XlsTextBoxShape shape = sheet.TextBoxes[0] as XlsTextBoxShape;

            //Extract text from the text box.
            StringBuilder content = new StringBuilder();
            content.AppendLine("The text extracted from the TextBox is: ");
            content.AppendLine(shape.Text);

            String result = "Result-ExtractTextFromATextbox.txt";

            //Save to file.
            File.WriteAllText(result, content.ToString());

            //Launch the file.
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
