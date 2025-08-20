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
        {// Create a new workbook
            Workbook workbook = new Workbook();

            // Clear all existing worksheets in the workbook
            workbook.Worksheets.Clear();

            // Create a new empty worksheet
            Worksheet sheet = workbook.CreateEmptySheet();

            // Create a StringBuilder to store the output text
            StringBuilder sb = new StringBuilder();

            // Get the default row count and column count of the worksheet
            int rowCount = sheet.Rows.Length;
            int columnCount = sheet.Columns.Length;

            // Append the row count and column count to the StringBuilder
            sb.AppendLine("The default row count is: " + rowCount);
            sb.AppendLine("The default column count is: " + columnCount);

            // Specify the output file name
            string output = "GetDefaultRowAndColumnCount.txt";

            // Write the contents of the StringBuilder to a text file
            File.WriteAllText(output, sb.ToString());

            // Dispose of the workbook object to release resources
            workbook.Dispose();

            // Launch the file
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
