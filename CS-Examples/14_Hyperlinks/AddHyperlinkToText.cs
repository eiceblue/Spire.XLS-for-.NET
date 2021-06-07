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
            //Create a workbook
			Workbook workbook = new Workbook();

            //Load the document from disk
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\CommonTemplate1.xlsx");

            //Get the first sheet
            Worksheet sheet = workbook.Worksheets[0];

            //Add url link
            HyperLink UrlLink = sheet.HyperLinks.Add(sheet.Range["D10"]);
            UrlLink.TextToDisplay = sheet.Range["D10"].Text;
            UrlLink.Type = HyperLinkType.Url;
            UrlLink.Address = "http://en.wikipedia.org/wiki/Chicago";

            //Add email link
            HyperLink MailLink = sheet.HyperLinks.Add(sheet.Range["E10"]);
            MailLink.TextToDisplay = sheet.Range["E10"].Text;
            MailLink.Type = HyperLinkType.Url;
            MailLink.Address = "mailto:Amor.Aqua@gmail.com";

            //Save the document
            string output = "AddHyperlinkToText.xlsx";
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
