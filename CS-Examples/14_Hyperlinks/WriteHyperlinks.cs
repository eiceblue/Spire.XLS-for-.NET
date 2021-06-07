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
			Workbook workbook = new Workbook();
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\WriteHyperlinks.xlsx");
			Worksheet sheet = workbook.Worksheets[0];

			sheet.Range["B9"].Text = "Home page";
			HyperLink hylink1 = sheet.HyperLinks.Add(sheet.Range["B10"]);
			hylink1.Type = HyperLinkType.Url;
			hylink1.Address = @"http://www.e-iceblue.com";

			sheet.Range["B11"].Text = "Support";
			HyperLink hylink2 = sheet.HyperLinks.Add(sheet.Range["B12"]);
			hylink2.Type = HyperLinkType.Url;
			hylink2.Address = "mailto:support@e-iceblue.com";

            sheet.Range["B13"].Text = "Forum";
            HyperLink hylink3 = sheet.HyperLinks.Add(sheet.Range["B14"]);
            hylink3.Type = HyperLinkType.Url;
            hylink3.Address = "https://www.e-iceblue.com/forum/";

            String result = "Output_WriteHyperlinks.xlsx";
            workbook.SaveToFile(result, ExcelVersion.Version2010);
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
