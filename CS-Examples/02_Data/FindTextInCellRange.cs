using System;
using System.Windows.Forms;
using Spire.Xls;
using System.Text;
using System.IO;

namespace FindTextInCellRange
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        private void btnRun_Click(object sender, System.EventArgs e)
        {   
            // Create a new Workbook object
            Workbook workbook = new Workbook();

            // Load the workbook from the specified file path 
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\FindTextFromRangeWithFindOptions.xlsx");

            // Get the first worksheet in the workbook
            Worksheet sheet = workbook.Worksheets[0];

            // Create a StringBuilder object to store the results
            StringBuilder builder = new StringBuilder();

            // Define the range to search for the text
            //CellRange range = sheet.Range[16, 1, 20, 2];
            CellRange range = sheet.Range["A16:B20"];

            // Find all occurrences of the specified text in the range
            CellRange[] resultRange = range.FindAll("e-iceblue1", FindType.Text, ExcelFindOptions.MatchEntireCellContent | ExcelFindOptions.MatchCase);

            // Check if any occurrences were found
            if (resultRange.Length != 0)
            {
                // Iterate through the found ranges and append their addresses to the StringBuilder
                foreach (CellRange r in resultRange)
                {
                    string address = r.RangeAddress;
                    builder.AppendLine("In the range 'A16:B20', the address of the cell containing 'e-iceblue1' is: " + address);
                }
            }

            // Define the output file path
            string result = "Result_out.txt";

            // Write the contents of the StringBuilder to the output file
            File.WriteAllText(result, builder.ToString());

            // Dispose the workbook object
            workbook.Dispose();

            // View the result TXT file
            OutputViewer(result);
        }
      
        private void OutputViewer(string fileName)
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
