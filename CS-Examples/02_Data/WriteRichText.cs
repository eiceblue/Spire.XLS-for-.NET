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
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Load the workbook from file
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\WriteRichText.xlsx");

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Create font styles for different formatting options
            ExcelFont fontBold = workbook.CreateFont();
            fontBold.IsBold = true;

            ExcelFont fontUnderline = workbook.CreateFont();
            fontUnderline.Underline = FontUnderlineType.Single;

            ExcelFont fontItalic = workbook.CreateFont();
            fontItalic.IsItalic = true;

            ExcelFont fontColor = workbook.CreateFont();
            fontColor.KnownColor = ExcelColors.Green;

            // Get the rich text object for cell B11 in the worksheet
            RichText richText = sheet.Range["B11"].RichText;

            // Set the text content for the rich text
            richText.Text = "Bold and underlined and italic and colored text.";

            // Apply different font styles to specific parts of the rich text
            richText.SetFont(0, 3, fontBold); 
            richText.SetFont(9, 18, fontUnderline);
            richText.SetFont(24, 29, fontItalic); 
            richText.SetFont(35, 41, fontColor); 

            // Save the modified workbook to the specified file in Excel 2013 format
            workbook.SaveToFile("WriteRichText_result.xlsx", ExcelVersion.Version2013);

            // Dispose of the workbook object to release resources
            workbook.Dispose();

            // View File
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
