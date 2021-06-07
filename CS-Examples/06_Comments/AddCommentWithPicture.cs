using System;
using System.Drawing;
using System.Windows.Forms;
using Spire.Xls;

namespace AddCommentWithPicture
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

            sheet.Range["C6"].Text = "E-iceblue";
            //Add the comment
            ExcelComment comment = sheet.Range["C6"].AddComment();
            //Load the image file
            Image image = Image.FromFile(@"..\..\..\..\..\..\Data\Logo.png");
            //Fill the comment with a customized background picture
            comment.Fill.CustomPicture(image, "logo.png");
            //Set the height and width of comment
            comment.Height = image.Height;
            comment.Width = image.Width;
            comment.Visible = true;

            //Save the document
            string output = "AddCommentWithPicture.xlsx";
			workbook.SaveToFile(output, ExcelVersion.Version2010);

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
