using System;
using System.Data.OleDb;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;

using Spire.Xls;

namespace WriteHyperlinks
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

            // Load an existing workbook from a file
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\WriteHyperlinks.xlsx");

            // Get the first worksheet in the workbook
            Worksheet sheet = workbook.Worksheets[0];

            // Set the text for cell B9 as "Home page"
            sheet.Range["B9"].Text = "Home page";

            // Add a hyperlink to cell B10
            HyperLink hylink1 = sheet.HyperLinks.Add(sheet.Range["B10"]);
            hylink1.Type = HyperLinkType.Url;
            hylink1.Address = @"http://www.e-iceblue.com";

            // Set the text for cell B11 as "Support"
            sheet.Range["B11"].Text = "Support";

            // Add a hyperlink to cell B12
            HyperLink hylink2 = sheet.HyperLinks.Add(sheet.Range["B12"]);
            hylink2.Type = HyperLinkType.Url;
            hylink2.Address = "mailto:support@e-iceblue.com";

            // Set the text for cell B13 as "Forum"
            sheet.Range["B13"].Text = "Forum";

            // Add a hyperlink to cell B14
            HyperLink hylink3 = sheet.HyperLinks.Add(sheet.Range["B14"]);
            hylink3.Type = HyperLinkType.Url;
            hylink3.Address = "https://www.e-iceblue.com/forum/";

            // Specify the output file name for the modified workbook
            String result = "Output_WriteHyperlinks.xlsx";

            // Save the modified workbook to the specified file using Excel 2010 format
            workbook.SaveToFile(result, ExcelVersion.Version2010);

            // Dispose of the workbook object to release resources
            workbook.Dispose();

            // Launch the file
            ExcelDocViewer(result);
		}

        private void btnClose_Click(object sender, EventArgs e)
        {
            Close();
        }
        private void ExcelDocViewer(string fileName)
        {
            try
            {
                System.Diagnostics.Process.Start(fileName);
            }
            catch { }
        }

	}
}
