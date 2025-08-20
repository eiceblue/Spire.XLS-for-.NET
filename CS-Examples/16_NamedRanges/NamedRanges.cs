using System;
using System.Data.OleDb;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;

using Spire.Xls;
using Spire.Xls.Core;

namespace NamedRanges
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }   
        private void btnRun_Click(object sender, System.EventArgs e)
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Load an existing workbook from a file
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\NamedRanges.xlsx");

            // Get the first worksheet in the workbook
            Worksheet sheet = workbook.Worksheets[0];

            // Create a new named range
            INamedRange NamedRange = workbook.NameRanges.Add("NewNamedRange");

            // Set the range of the named range to cover cells A8 to E12 on the worksheet
            NamedRange.RefersToRange = sheet.Range["A8:E12"];

            // Specify the output file name for the modified workbook
            string result = "NamedRanges_result.xlsx";

            // Save the modified workbook to the specified file using Excel 2013 format
            workbook.SaveToFile(result, ExcelVersion.Version2013);

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
