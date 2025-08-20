using System;
using System.Windows.Forms;
using Spire.Xls;
using System.Text;
using System.IO;

namespace GetPageCount
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

            // Get Split Page Info
            var pageInfoList = workbook.GetSplitPageInfo();
            StringBuilder sb = new StringBuilder();

            // Get page count
            for (int i = 0; i < workbook.Worksheets.Count; i++)
            {
                string sheetname = workbook.Worksheets[i].Name;
                int pagecount = pageInfoList[i].Count;
                sb.AppendLine(sheetname + "'s page count is: " + pagecount);
            }

            // Specify the output file path and name
            string output = "GetPageCount.txt";

            // Save a text file
            File.WriteAllText(output, sb.ToString());

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
