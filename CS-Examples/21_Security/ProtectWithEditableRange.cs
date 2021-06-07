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
            //Create a workbook and load a file
            Workbook workbook = new Workbook();
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\ProtectWithEditableRange.xlsx");

            //Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            //Define the specified ranges to allow users to edit while sheet is protected
            sheet.AddAllowEditRange("EditableRanges", sheet.Range["B4:E12"]);

            //Protect worksheet with a password.
            sheet.Protect("TestPassword", SheetProtectionType.All);

            String result = "ProtectWithEditableRange_result.xlsx";
            //Save the document and launch it
            workbook.SaveToFile(result, ExcelVersion.Version2010);
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
