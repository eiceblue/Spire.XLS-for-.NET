using System;
using System.Windows.Forms;
using Spire.Xls;
using Spire.Xls.Collections;

namespace ModifyHyperlink
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
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\ModifyHyperlink.xlsx");

            //Get the collection of all hyperlinks in the worksheet
            Worksheet sheet = workbook.Worksheets[0];

            //Change the values of TextToDisplay and Address property 
            HyperLinksCollection links = sheet.HyperLinks;
            links[0].TextToDisplay = "Spire.XLS for .NET";
            links[0].Address = "http://www.e-iceblue.com/Introduce/excel-for-net-introduce.html";

            //Save the document
            string output = "ModifyHyperlinkResult.xlsx";
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
