using Spire.Xls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace CopyVisibleSheets
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

            // Load a CSV file into the workbook
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\CopyVisibleSheets.xlsx");

            // Create a new workbook to copy visible sheets
            Workbook workbookNew = new Workbook();
            workbookNew.Version = ExcelVersion.Version2013;
            workbookNew.Worksheets.Clear();

            // Loop through the worksheets in the original workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Check if the worksheet is visible
                if (sheet.Visibility == WorksheetVisibility.Visible)
                {
                    // Copy the visible sheet to the new workbook
                    string name = sheet.Name;
                    workbookNew.Worksheets.AddCopy(sheet);
                }
            }

            // Save the new workbook with copied visible sheets
            string result = "CopyVisibleSheets_out.xlsx";
            workbookNew.SaveToFile(result, ExcelVersion.Version2013);

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
        private void btnClose_Click_1(object sender, EventArgs e)
        {
            Close();
        }
    }
}
