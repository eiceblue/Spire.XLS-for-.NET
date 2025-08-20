using System;
using System.Windows.Forms;
using Spire.Xls;

namespace ToOfficeOpenXML
{
	public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
		private void btnRun_Click(object sender, System.EventArgs e)
		{
            // Create a workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet from the workbook
            Worksheet sheet = workbook.Worksheets[0];

            // Set the text "Hello World" in cell A1 of the worksheet
            sheet.Range["A1"].Text = "Hello World";

            // Apply the color Gray25Percent to cell B1 using a known color
            sheet.Range["B1"].Style.KnownColor = ExcelColors.Gray25Percent;

            // Apply the color Gold to cell C1 using a known color
            sheet.Range["C1"].Style.KnownColor = ExcelColors.Gold;

            // Save the workbook as an XML file with the name "sample.xml"
            workbook.SaveAsXml("sample.xml");

            // Dispose of the workbook object to release resources
            workbook.Dispose();

            // Launch the file
            ExcelDocViewer("sample.xml");

            
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
