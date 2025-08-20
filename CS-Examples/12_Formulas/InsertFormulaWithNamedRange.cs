using System;
using System.Data.OleDb;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;

using Spire.Xls;
using Spire.Xls.Core;


namespace InsertFormulaWithNamedRange
{
    public partial class Form1 : Form
    {

        public Form1()
        {
            InitializeComponent();
        }

        private void btnRun_Click(object sender, System.EventArgs e)
        {
            // Create a workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Set values for cells A1 and A2
            sheet.Range["A1"].Value = "1";
            sheet.Range["A2"].Value = "1";

            // Create a named range
            INamedRange namedRange = workbook.NameRanges.Add("NewNamedRange");

            // Set the local name and formula for the named range
            namedRange.NameLocal = "=SUM(A1+A2)";

            // Set the formula for cell C1 to reference the named range
            sheet.Range["C1"].Formula = "NewNamedRange";

            // Save the workbook to the specified file in Excel 2010 format
            string result = "result.xlsx";
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
