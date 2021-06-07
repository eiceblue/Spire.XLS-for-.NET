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

            Worksheet sheet = workbook.Worksheets[0];

            //Add the description text
            sheet.Columns[0].ColumnWidth = 22;
            sheet.Range["A1"].Text = "Image Hyperlink";
            sheet.Range["A1"].Style.VerticalAlignment = VerticalAlignType.Top;

            //Insert an image to a specific cell
            string picPath = @"..\..\..\..\..\..\Data\SpireXls.png";
            ExcelPicture picture = sheet.Pictures.Add(2, 1, picPath);
            //Add a hyperlink to the image
            picture.SetHyperLink("https://www.e-iceblue.com/Introduce/excel-for-net-introduce.html", true);

            //Save the document
            string output = "AddImageHyperlink.xlsx";
			workbook.SaveToFile(output, ExcelVersion.Version2013);

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
