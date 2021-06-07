using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;

using Spire.Xls;
using System.Text;
using System.IO;

namespace FindFormulaCells
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

            //Find the cells that contain formula "=SUM(A11,A12)"
            CellRange[] ranges = sheet.FindAll("=SUM(A11,A12)", FindType.Formula, ExcelFindOptions.None);

            //Create a string builder
            StringBuilder builder = new StringBuilder();

            //Append the address of found cells to builder
            if (ranges.Length != 0)
            {
                foreach (CellRange range in ranges)
                {
                    string address = range.RangeAddress;
                    builder.AppendLine("The address of found cell is: " + address);
                }
            }
            else
            {
                builder.AppendLine("No cell contain the formula");
            }

            //Save to txt file
            string result = "FindFormulaCells_out.txt";
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
