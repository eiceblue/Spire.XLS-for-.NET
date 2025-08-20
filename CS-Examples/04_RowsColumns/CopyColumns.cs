using Spire.Xls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Xls.Core.Spreadsheet.PivotTables;

namespace CopyColumns
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

            // Load an existing workbook with a pivot table from a file
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\Copying.xls");

            // Get the first worksheet in the workbook
            Worksheet sheet1 = workbook.Worksheets[0];

            // Get the second worksheet in the workbook
            Worksheet sheet2 = workbook.Worksheets[1];

            // Copy the first column (column index 0) to the third column (column index 2) in the same sheet
            sheet1.Copy(sheet1.Columns[0], sheet1.Columns[2], true, true, true);

            // Copy the first column (column index 0) to the second column (column index 1) in a different sheet
            sheet1.Copy(sheet1.Columns[0], sheet2.Columns[1], true, true, true);

            // Specify the output file name for the result
            string result = "CopyColumns_result.xlsx";

            // Save the modified workbook to the specified file using Excel 2010 format
            workbook.SaveToFile(result, ExcelVersion.Version2010);

            // Dispose of the workbook object to release resources
            workbook.Dispose();

            //View the document
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
