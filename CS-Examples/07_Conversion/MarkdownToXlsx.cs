using System;
using System.Windows.Forms;
using Spire.Xls;

namespace MarkdownToXlsx
{
	public partial class Form1 : Form
	{
        public Form1()
        {
            InitializeComponent();
        }
		private void btnRun_Click(object sender, System.EventArgs e)
		{
            // Create a new Workbook instance
            Workbook workbook = new Workbook();

            // Load content from a Markdown file into the workbook
            workbook.LoadFromMarkdown(@"..\..\..\..\..\..\Data\sample.md");

            // Define the output file name for the saved Excel file
            String result = "MarkdownToXlsx.xlsx";

            // Save the workbook to a file in Excel 2016 format (.xlsx)
            workbook.SaveToFile(result, ExcelVersion.Version2016);

            // Release the resources used by the workbook object
            workbook.Dispose();

            // Launch the file
            ExcelDocViewer(result);	
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
