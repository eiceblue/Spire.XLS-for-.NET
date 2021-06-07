using System;
using System.Windows.Forms;
using Spire.Xls;
using System.Text;
using System.IO;

namespace GetWorksheetNames
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
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\WorksheetSample3.xlsx");

            //Get the names of all worksheets
            StringBuilder sb = new StringBuilder();
            foreach(Worksheet sheet in workbook.Worksheets)
            {
                sb.AppendLine(sheet.Name);
            }

            //Save to the Text file
            string output = "GetWorksheetNames.txt";
            File.WriteAllText(output, sb.ToString());

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
