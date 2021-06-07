using System;
using System.Data.OleDb;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;

using Spire.Xls;

namespace WriteComment
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
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\WriteComment.xlsx");
			Worksheet sheet = workbook.Worksheets[0];

			//Creates font
            ExcelFont font=workbook.CreateFont();
            font.FontName="Arial";
            font.Size=11;
            font.KnownColor = ExcelColors.Orange;
			ExcelFont fontBlue = workbook.CreateFont();
			fontBlue.KnownColor = ExcelColors.LightBlue;
			ExcelFont fontGreen = workbook.CreateFont();
			fontGreen.KnownColor = ExcelColors.LightGreen;

			CellRange range = sheet.Range["B11"];
			range.Text = "Regular comment";
			range.Comment.Text = "Regular comment";
            range.AutoFitColumns();
			//Regular comment
          

            range = sheet.Range["B12"];
			range.Text = "Rich text comment";
            range.RichText.SetFont(0, 16, font);
            range.AutoFitColumns();
			//Rich text comment
			range.Comment.RichText.Text = "Rich text comment";
			range.Comment.RichText.SetFont(0,4, fontGreen);
			range.Comment.RichText.SetFont(5,9, fontBlue);

            string result = "WriteComment_result.xlsx";
            workbook.SaveToFile(result, ExcelVersion.Version2007);
            ExcelDocViewer(result);
		}
		private void ExcelDocViewer( string fileName )
		{
			try
			{
				System.Diagnostics.Process.Start(fileName);
			}
			catch{}
		}

        private void btnClose_Click(object sender, EventArgs e)
        {
            Close();
        }

	}
}
