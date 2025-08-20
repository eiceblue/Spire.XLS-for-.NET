using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using Spire.Xls;
using System.Text;
using System.IO;

namespace FindDataInSpecificRange
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        private void btnRun_Click(object sender, System.EventArgs e)
        {
            //Create a workbook
            Workbook workbook = new Workbook();

            //Load the document from disk
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\FindCellsSample.xlsx");

            //Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            //Specify a range
            CellRange range = sheet.Range[1, 1, 12, 8];

            //Create a string builder
            StringBuilder builder = new StringBuilder();

            //Find text from this range
            FindTextFromRange(range, builder);

            //Find number from this range
            FindNumberFromRange(range, builder);

            // Specify the name for the resulting Excel file
            string result = "FindDataInSpecificRange_out.txt";

            // Save to text
            File.WriteAllText(result, builder.ToString());

            // Dispose of the workbook object to free up resources
            workbook.Dispose();

            //Launch the file
            OutputViewer(result);
        }
        private void FindTextFromRange(CellRange range, StringBuilder builder)
        {
            //Find string from this range
            CellRange[] textRanges = range.FindAllString("E-iceblue", false, false);

            //Append the address of found cells in builder
            if (textRanges.Length != 0)
            {
                foreach (CellRange r in textRanges)
                {
                    string address = r.RangeAddress;
                    builder.AppendLine("The address of found text cell is: " + address);
                }
            }
            else
            {
                builder.AppendLine("No cell contain the text");
            }
        }
        private void FindNumberFromRange(CellRange range, StringBuilder builder)
        {
            //Find number from this range
            CellRange[] numberRanges = range.FindAllNumber(100, true);

            //Append the address of found cells in builder
            if (numberRanges.Length != 0)
            {
                foreach (CellRange r in numberRanges)
                {
                    string address = r.RangeAddress;
                    builder.AppendLine("The address of found number cell is: " + address);
                }
            }
            else
            {
                builder.AppendLine("No cell contain the number");
            }
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
