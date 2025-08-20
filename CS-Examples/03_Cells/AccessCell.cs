using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using Spire.Xls;
using System.Text;
using System.IO;

namespace AccessCell
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
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\AccessCell.xlsx");

            StringBuilder builder = new StringBuilder();

            //Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            //Access cell by its name
            CellRange range1 = sheet.Range["A1"];
            builder.AppendLine("Value of range1: " + range1.Text);

            //Access cell by index of row and column
            CellRange range2 = sheet.Range[2,1];
            builder.AppendLine("Value of range2: " + range2.Text);

            //Access cell in cell collection
            CellRange range3 = sheet.Cells[2];
            builder.AppendLine("Value of range3: " + range3.Text);

            //Specify the filename for the resulting file
            string result = "AccessCell_out.txt";

            // Save to text file
            File.WriteAllText(result, builder.ToString());

            // Dispose of the workbook object to release resources
            workbook.Dispose();

            //Launch the txt file
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
