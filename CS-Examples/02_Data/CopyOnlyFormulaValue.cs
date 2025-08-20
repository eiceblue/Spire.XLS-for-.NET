using Spire.Xls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace CopyOnlyFormulaValue
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Load an existing workbook from a file
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\CopyOnlyFormulaValue.xlsx");

            // Get the first worksheet in the workbook
            Worksheet sheet = workbook.Worksheets[0];

            // Set the copy option to only copy the formula values
            CopyRangeOptions copyOptions = CopyRangeOptions.OnlyCopyFormulaValue;

            // Copy a range of cells from A2:C2 to A5:C5 using the specified copy options
            sheet.Copy(sheet.Range["A2:C2"], sheet.Range["A5:C5"], copyOptions);

            // Save the modified workbook to a new file named "result.xlsx" in Excel 2010 format
            workbook.SaveToFile("result.xlsx", ExcelVersion.Version2010);

            // Dispose of the workbook object to free up resources
            workbook.Dispose();

            //View the document
            FileViewer("result.xlsx");
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
