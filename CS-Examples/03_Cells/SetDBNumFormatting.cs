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
            //Create a workbook
            Workbook workbook = new Workbook();

            workbook.CreateEmptySheets(1);

            //Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            //Set value for cells
            sheet.Range["A1"].Value2 = 123;
            sheet.Range["A2"].Value2 = 456;
            sheet.Range["A3"].Value2 = 789;

            //Get the cell range
            CellRange range = sheet.Range["A1:A3"];

            //Set the DB num format
            range.NumberFormat = "[DBNum2][$-804]General";

            //Auto fit columns
            range.AutoFitColumns();

            //Save the document
            string output = "SetDBNumFormatting_out.xlsx";
            workbook.SaveToFile(output, ExcelVersion.Version2013);

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
