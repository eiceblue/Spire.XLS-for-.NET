using System;
using System.Data.OleDb;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;

using Spire.Xls;

namespace WriteRichText
{
	public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

		private void btnRun_Click(object sender, System.EventArgs e)
		{
			Workbook workbook = new Workbook();
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\WriteRichText.xlsx");
			Worksheet sheet = workbook.Worksheets[0];

			ExcelFont fontBold = workbook.CreateFont();
			fontBold.IsBold = true;

			ExcelFont fontUnderline = workbook.CreateFont();
			fontUnderline.Underline = FontUnderlineType.Single;

            ExcelFont fontItalic = workbook.CreateFont();
            fontItalic.IsItalic = true;

			ExcelFont fontColor = workbook.CreateFont();
			fontColor.KnownColor  = ExcelColors.Green; 

			RichText richText = sheet.Range["B11"].RichText;
            richText.Text = "Bold and underlined and italic and colored text.";
			richText.SetFont(0,3,fontBold);
			richText.SetFont(9,18,fontUnderline);
            richText.SetFont(24, 29, fontItalic);
			richText.SetFont(35,41,fontColor);

            workbook.SaveToFile("WriteRichText_result.xlsx",ExcelVersion.Version2013);
            ExcelDocViewer("WriteRichText_result.xlsx");
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
