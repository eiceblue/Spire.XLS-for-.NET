using System;
using System.Data.OleDb;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;

using Spire.Xls;
using Spire.Xls.Charts;

namespace ProtectCell
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

            // Load file from disk
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\ProtectCell.xlsx");

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Protect cell
            sheet.Range["B3"].Style.Locked = true;
            sheet.Range["C3"].Style.Locked = false;

            // Set password
            sheet.Protect("TestPassword", SheetProtectionType.All);

            // Specify the output filename for the workbook
            String result = "ProtectCell_result.xlsx";

            // Save the modified workbook to a file (in Excel 2013 format)
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
