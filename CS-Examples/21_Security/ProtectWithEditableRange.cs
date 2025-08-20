using Spire.Xls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace ProtectWithEditableRange
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

            // Load an existing Excel document from file
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\ProtectWithEditableRange.xlsx");

            // Get the first worksheet from the workbook
            Worksheet sheet = workbook.Worksheets[0];

            // Define the specified ranges that allow users to edit while the sheet is protected
            sheet.AddAllowEditRange("EditableRanges", sheet.Range["B4:E12"]);

            // Protect the worksheet with a password
            sheet.Protect("TestPassword", SheetProtectionType.All);

            // Specify the output filename for the workbook
            String result = "ProtectWithEditableRange_result.xlsx";

            // Save the modified workbook to a file
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
        private void btnAbout_Click(object sender, EventArgs e)
        {
            Close();
        }
    }
}
