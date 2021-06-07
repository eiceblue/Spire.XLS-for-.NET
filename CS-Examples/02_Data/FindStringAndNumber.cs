using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using Spire.Xls;
using System.Text;
using System.IO;

namespace FindStringAndNumber
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

            //Find cells with the input string
            CellRange[] textRanges = sheet.FindAllString("E-iceblue", false, false);

            //Create a string builder
            StringBuilder builder = new StringBuilder();

            //Append the address of found cells in builder
            if (textRanges.Length != 0)
            {
                foreach (CellRange range in textRanges)
                {
                    string address = range.RangeAddress;
                    builder.AppendLine("The address of found text cell is: " + address);
                }
            }
            else
            {
                builder.AppendLine("No cells that contain the text");
            }

            //Find cells with the input integer or double
            CellRange[] numberRanges = sheet.FindAllNumber(100, true);

            //Append the address of found cells in builder
            if (numberRanges.Length != 0)
            {
                foreach (CellRange range in numberRanges)
                {
                    string address = range.RangeAddress;
                    builder.AppendLine("The address of found number cell is: " + address);
                }
            }
            else
            {
                builder.AppendLine("No cells that contain the number");
            }

            //Save to txt file
            string result = "FindStringAndNumber_out.txt";
            File.WriteAllText(result, builder.ToString());

            //Launch the file
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
