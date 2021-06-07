using System;
using System.Data.OleDb;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using Spire.Xls;
using Spire.Xls.Core.Spreadsheet;

namespace AlignPictureWithinCell
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

            sheet.Range["A1"].Text = "Align Picture Within A Cell:";
            sheet.Range["A1"].Style.VerticalAlignment = VerticalAlignType.Top;

            //Insert an image to the specific cell.
            string picPath = @"..\..\..\..\..\..\Data\SpireXls.png";
            ExcelPicture picture = sheet.Pictures.Add(1, 1, picPath);

            //Adjust the column width and row height so that the cell can contain the picture.
            sheet.Columns[0].ColumnWidth = 40;
            sheet.Rows[0].RowHeight = 200;

            //Vertically and horizontally align the image.
            picture.LeftColumnOffset = 100;
            picture.TopRowOffset = 25;

            String result = "Result-AlignPictureWithinCell.xlsx";

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
