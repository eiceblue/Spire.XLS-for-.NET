using Spire.Xls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace RemoveDuplicatedRows
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
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\DuplicatedRows.xlsx");

            // Get the first worksheet in the workbook
            Worksheet sheet = workbook.Worksheets[0];

            // Remove duplicated rows in the worksheet
            sheet.RemoveDuplicates();

            // Remove the duplicate rows within the specified range
            // sheet.RemoveDuplicates(int startRow, int startColumn, int endRow, int endColumn);
            // Remove the duplicated rows based on specific columns and headers
            // sheet.RemoveDuplicates(int startRow, int startColumn, int endRow, int endColumn, boolean hasHeaders, int[] columnOffsets)

            // Specify the output file name for the result
            string result = "RemoveDuplicatedRows_result.xlsx";

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
