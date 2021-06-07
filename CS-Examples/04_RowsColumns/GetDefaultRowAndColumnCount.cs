using System;
using System.IO;
using System.Text;
using System.Windows.Forms;
using Spire.Xls;

namespace GetDefaultRowAndColumnCount
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
            //Clear all worksheets
            workbook.Worksheets.Clear();

            //Create a new worksheet
            Worksheet sheet = workbook.CreateEmptySheet();
            StringBuilder sb = new StringBuilder();
            //Get row and column count
            int rowCount = sheet.Rows.Length;
            int columnCount = sheet.Columns.Length;

            sb.AppendLine("The default row count is :" + rowCount);
            sb.AppendLine("The default column count is :" + columnCount);

            //Save to Text file
            string output = "GetDefaultRowAndColumnCount.txt";
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
