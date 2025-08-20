using Spire.Xls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace DaysFormula
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            // Create a new workbook object
            Workbook workbook = new Workbook();

            // Load an existing Excel document
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\Template_Xls_12.xlsx");

            // Get the first sheet from the workbook
            Worksheet sheet = workbook.Worksheets[0];

            // Add a formula to cell C4
            sheet.Range["C4"].Formula = "=DAYS(A8,A1)";

            // Calculate all values in the workbook
            workbook.CalculateAllValue();

            // Specify the name for the resulting Excel file
            String result = "DaysFormula_result.xlsx";

            // Save the modified workbook to a file
            workbook.SaveToFile(result, ExcelVersion.Version2016);

            // Dispose of the workbook object to release resources
            workbook.Dispose();

            // View the document
            FileViewer(result);
        }

        private void FileViewer(string fileName)
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
