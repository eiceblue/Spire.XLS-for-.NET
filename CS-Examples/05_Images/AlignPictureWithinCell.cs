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
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet in the workbook
            Worksheet sheet = workbook.Worksheets[0];

            // Set the text in cell A1 as "Align Picture Within A Cell:"
            sheet.Range["A1"].Text = "Align Picture Within A Cell:";

            // Set the vertical alignment of cell A1 to top
            sheet.Range["A1"].Style.VerticalAlignment = VerticalAlignType.Top;

            // Insert an image at the specific cell (1, 1)
            string picPath = @"..\..\..\..\..\..\Data\SpireXls.png";
            ExcelPicture picture = sheet.Pictures.Add(1, 1, picPath);

            // Adjust the column width and row height so that the cell can contain the picture
            sheet.Columns[0].ColumnWidth = 40;
            sheet.Rows[0].RowHeight = 200;

            // Set the horizontal offset of the image within the cell to 100
            picture.LeftColumnOffset = 100;

            // Set the vertical offset of the image within the cell to 25
            picture.TopRowOffset = 25;

            // Specify the name of the resulting Excel file
            string result = "Result-AlignPictureWithinCell.xlsx";

            // Save the workbook to the specified file using Excel 2013 format
            workbook.SaveToFile(result, ExcelVersion.Version2013);

            // Dispose of the workbook object to release resources
            workbook.Dispose();

            // Launch the file.
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
