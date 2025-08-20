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
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Load an existing workbook from a file
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\ModifyHyperlink.xlsx");

            // Get the first worksheet in the workbook
            Worksheet sheet = workbook.Worksheets[0];

            // Get the collection of all hyperlinks in the worksheet
            HyperLinksCollection links = sheet.HyperLinks;

            // Modify the values of TextToDisplay and Address properties of the first hyperlink
            links[0].TextToDisplay = "Spire.XLS for .NET";
            links[0].Address = "http://www.e-iceblue.com/Introduce/excel-for-net-introduce.html";

            // Specify the output file name for the modified workbook
            string output = "ModifyHyperlinkResult.xlsx";

            // Save the modified workbook to the specified file using Excel 2013 format
            workbook.SaveToFile(output, ExcelVersion.Version2013);

            // Dispose of the workbook object to release resources
            workbook.Dispose();

            // Launch the Excel file
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
