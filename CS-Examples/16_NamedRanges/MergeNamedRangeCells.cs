using System;
using System.Data.OleDb;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;

using Spire.Xls;
using Spire.Xls.Core;


namespace MergeNamedRangeCells
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
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\AllNamedRanges.xlsx");

            // Get a specific named range by its index
            INamedRange NamedRange = workbook.NameRanges[0];

            // Get the range of the named range
            IXLSRange range = NamedRange.RefersToRange;

            // Merge cells within the range
            range.Merge();

            // Specify the output file name for the result
            string result = "result.xlsx";

            // Save the modified workbook to the specified file using Excel 2010 format
            workbook.SaveToFile(result, ExcelVersion.Version2010);

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
