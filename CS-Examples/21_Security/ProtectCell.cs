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
            //Create a workbook and load a file
            Workbook workbook = new Workbook();
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\ProtectCell.xlsx");

            //Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            //Protect cell
            sheet.Range["B3"].Style.Locked = true;
            sheet.Range["C3"].Style.Locked = false;

            sheet.Protect("TestPassword", SheetProtectionType.All);

            String result = "ProtectCell_result.xlsx";
			//Save the document and launch it
            workbook.SaveToFile(result, ExcelVersion.Version2013);
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
