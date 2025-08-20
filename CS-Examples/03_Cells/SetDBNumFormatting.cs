using System;
using System.Data.OleDb;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;

using Spire.Xls;

namespace SetDBNumFormatting
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

            // Create an empty worksheet
            workbook.CreateEmptySheets(1);

            // Get the first worksheet in the workbook
            Worksheet sheet = workbook.Worksheets[0];

            // Set values for cells A1, A2, and A3
            sheet.Range["A1"].Value2 = 123;
            sheet.Range["A2"].Value2 = 456;
            sheet.Range["A3"].Value2 = 789;

            // Get the cell range A1:A3
            CellRange range = sheet.Range["A1:A3"];

            // Set the DB num format for the range
            range.NumberFormat = "[DBNum2][$-804]General";

            // Auto fit columns for the range
            range.AutoFitColumns();

            // Save the modified workbook to a file
            string output = "SetDBNumFormatting_out.xlsx";
            workbook.SaveToFile(output, ExcelVersion.Version2013);

            // Dispose of the workbook object to release resources
            workbook.Dispose();

            //Launch the Excel file
            ExcelDocViewer(output);
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
