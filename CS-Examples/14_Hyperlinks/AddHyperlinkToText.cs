using System;
using System.Windows.Forms;
using Spire.Xls;

namespace AddHyperlinkToText
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

            // Load the document from disk
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\CommonTemplate1.xlsx");

            // Get the first sheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add url link
            HyperLink UrlLink = sheet.HyperLinks.Add(sheet.Range["D10"]);
            // Set display text
            UrlLink.TextToDisplay = sheet.Range["D10"].Text;
            // Set url link type
            UrlLink.Type = HyperLinkType.Url;
            // Set url address
            UrlLink.Address = "http://en.wikipedia.org/wiki/Chicago";

            //Add email link
            HyperLink MailLink = sheet.HyperLinks.Add(sheet.Range["E10"]);
            // Set display text
            MailLink.TextToDisplay = sheet.Range["E10"].Text;
            // Set mail link type
            MailLink.Type = HyperLinkType.Url;
            // Set mail address
            MailLink.Address = "mailto:Amor.Aqua@gmail.com";

            // Specify the file name for the resulting Excel file
            string output = "AddHyperlinkToText.xlsx";

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
