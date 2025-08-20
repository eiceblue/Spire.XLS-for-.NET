using System;
using System.IO;
using System.Text;
using System.Windows.Forms;
using Spire.Xls;

namespace GetPaperSize
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
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\WorksheetSample2.xlsx");

            StringBuilder sb = new StringBuilder();
            

            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Get page widths
                double width = sheet.PageSetup.PageWidth;

                // Get page height
                double height = sheet.PageSetup.PageHeight;
                sb.AppendLine(sheet.Name);
                sb.AppendLine("Width: " + width + "\tHeight: " + height);
                sb.AppendLine();
            }

            //Save to Text file
            string output = "GetPaperSize.txt";
            File.WriteAllText(output, sb.ToString());

            //Launch the file
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
