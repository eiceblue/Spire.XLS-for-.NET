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
            // Create a workbook
			Workbook workbook = new Workbook();
            
            // Load file from disk
            Worksheet sheet = workbook.Worksheets[0];

            // Set value for the range
            sheet.Range["C6"].Text = "E-iceblue";

            // Add the comment
            ExcelComment comment = sheet.Range["C6"].AddComment();

            // Load the image file
            Image image = Image.FromFile(@"..\..\..\..\..\..\Data\Logo.png");

            // Fill the comment with a customized background picture
            comment.Fill.CustomPicture(image, "logo.png");

            // Set the height and width of comment
            comment.Height = image.Height;
            comment.Width = image.Width;
            comment.Visible = true;

            // Specify the resulting file name.
            string output = "AddCommentWithPicture.xlsx";

            // Save the modified workbook to a file using Excel 201 format.
            workbook.SaveToFile(output, ExcelVersion.Version2010);

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
