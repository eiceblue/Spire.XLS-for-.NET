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

namespace SetFontAndBackground
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

            //Get the textbox which will be edited.
            XlsTextBoxShape shape = sheet.TextBoxes[0] as XlsTextBoxShape;

            //Set the font and background color for the textbox.
            //Set font.
            ExcelFont font = workbook.CreateFont();
            //font.IsStrikethrough = true;
            font.FontName = "Century Gothic";
            font.Size = 10;
            font.IsBold = true;
            font.Color = Color.Blue;
            (new RichText(shape.RichText)).SetFont(0, shape.Text.Length - 1, font);

            //Set background color
            shape.Fill.FillType = ShapeFillType.SolidColor;
            shape.Fill.ForeKnownColor = ExcelColors.BlueGray;

            String result = "Result-SetFontAndBackgroundForTextBox.xlsx";

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
