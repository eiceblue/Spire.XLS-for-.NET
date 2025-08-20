using System;
using System.Windows.Forms;
using Spire.Xls;
using System.IO;

namespace FindTextByRegex
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnRun_Click(object sender, System.EventArgs e)
        {
            // Load an existing workbook from a file
            Workbook workbook = new Workbook();
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\FindTextByRegex.xlsx");

            // Get the first sheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Find cell ranges by Regex
            CellRange[] ranges = worksheet.FindAllString(".*North.", false, false, true);
            string information = "";

            // Get the information of every cell range
            foreach (CellRange range in ranges)
            {
                information += "RangeAddressLocal:" + range.RangeAddressLocal + "\r\n";
                information += "Text:" + range.Text + "\r\n";
            }

            // Specify the output file name for the result
            string result = "FindTextByRegex_result.txt";

            File.WriteAllText(result, information);

            // Dispose of the workbook object to release resources
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
