using System;
using System.Data.OleDb;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;

using Spire.Xls;
using Spire.Xls.Core;
using System.Text;
using System.IO;

namespace GetAllNamedRange
{
    public partial class Form1 : Form
    {

        public Form1()
        {
            InitializeComponent();
        }

        private void btnRun_Click(object sender, System.EventArgs e)
        {
            // Create a StringBuilder to store the result
            StringBuilder sb = new StringBuilder();

            // Load an existing workbook from a file
            Workbook workbook = new Workbook();
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\AllNamedRanges.xlsx");

            // Get all named ranges in the workbook
            INameRanges ranges = workbook.NameRanges;

            // Iterate over each named range
            foreach (INamedRange nameRange in ranges)
            {
                // Append the name of the current named range to the StringBuilder
                sb.Append(nameRange.Name + "\r\n");
            }

            // Specify the output file name for the result
            string result = "result.txt";

            // Write the content of the StringBuilder to the result file
            File.WriteAllText(result, sb.ToString());

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
