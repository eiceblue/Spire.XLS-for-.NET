using System;
using System.Windows.Forms;
using Spire.Xls;

namespace AddImageHyperlink
{

	public partial class Form1 : Form
	{
        public Form1()
        {
            InitializeComponent();
        }

		private void btnRun_Click(object sender, EventArgs e)
		{
            //Create a workbook
			Workbook workbook = new Workbook();

            // Load a Workbook from disk
            Worksheet sheet = workbook.Worksheets[0];

            // Set width for the first cloumn
            sheet.Columns[0].ColumnWidth = 22;
            // Set value for cell "A1"
            sheet.Range["A1"].Text = "Image Hyperlink";
            // Ser vertical alignment as top
            sheet.Range["A1"].Style.VerticalAlignment = VerticalAlignType.Top;

            // Insert an image to a specific cell
            string picPath = @"..\..\..\..\..\..\Data\SpireXls.png";
            ExcelPicture picture = sheet.Pictures.Add(2, 1, picPath);

            // Add a hyperlink to the image
            picture.SetHyperLink("https://www.e-iceblue.com/Introduce/excel-for-net-introduce.html", true);

            // Specify the file name for the resulting Excel file
            string output = "AddImageHyperlink.xlsx";

            // Save the workbook to the specified file in Excel 2010 format
            workbook.SaveToFile(output, ExcelVersion.Version2013);

            // Dispose of the workbook object to release resources
            workbook.Dispose();

            //Launch the Excel file
            ExcelDocViewer(output);
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
