using System;
using System.Drawing;
using System.Windows.Forms;
using Spire.Xls;

namespace SetPositionAndAlignment
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

            //Get the default first  worksheet 
            Worksheet sheet = workbook.Worksheets[0];

            // Create two font styles which will be used in comments
            ExcelFont font1 = workbook.CreateFont();
            font1.FontName = "Calibri";
            font1.Color = Color.Firebrick;
            font1.IsBold = true;
            font1.Size = 12;
            ExcelFont font2 = workbook.CreateFont();
            font2.FontName = "Calibri";
            font2.Color = Color.Blue;
            font2.Size = 12;
            font2.IsBold = true;

            // Add comment 1 and set its size, text, position and alignment
            sheet.Range["G5"].Text = "Spire.XLS";
            ExcelComment Comment1 = sheet.Range["G5"].Comment;
            Comment1.IsVisible = true;
            Comment1.Height = 150;
            Comment1.Width = 300;
            Comment1.RichText.Text = "Spire.XLS for .Net:\nStandalone Excel component to meet your needs for conversion, data manipulation, charts in workbook etc. ";
            Comment1.RichText.SetFont(0, 19, font1);
            Comment1.TextRotation = TextRotationType.LeftToRight;

            // Set the position of Comment
            Comment1.Top = 20;
            Comment1.Left = 40;

            // Set the alignment of text in Comment
            Comment1.VAlignment = CommentVAlignType.Center;
            Comment1.HAlignment = CommentHAlignType.Justified;

            // Add comment2 and set its size, text, position and alignment for comparison
            sheet.Range["D14"].Text = "E-iceblue";
            ExcelComment Comment2 = sheet.Range["D14"].Comment;
            Comment2.IsVisible = true;
            Comment2.Height = 150;
            Comment2.Width = 300;
            Comment2.RichText.Text = "About E-iceblue: \nWe focus on providing excellent office components for developers to operate Word, Excel, PDF, and PowerPoint documents.";
            Comment2.TextRotation = TextRotationType.LeftToRight;
            Comment2.RichText.SetFont(0, 16, font2);

            // Set the position of Comment
            Comment2.Top = 170;
            Comment2.Left = 450;

            // Set the alignment of text in Comment
            Comment2.VAlignment = CommentVAlignType.Top;
            Comment2.HAlignment = CommentHAlignType.Justified;

            // Save the document
            string output = "SetPositionAndAlignment.xlsx";
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
