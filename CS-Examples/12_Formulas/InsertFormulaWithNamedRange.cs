using System;
using System.Data.OleDb;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;

using Spire.Xls;
using Spire.Xls.Core;


namespace InsertFormulaWithNamedRange
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
            Worksheet sheet = workbook.Worksheets[0];

            //Set value
            sheet.Range["A1"].Value = "1";
            sheet.Range["A2"].Value = "1";

            //Create a named range
            INamedRange NamedRange = workbook.NameRanges.Add("NewNamedRange");

            NamedRange.NameLocal = "=SUM(A1+A2)";

            //Set the formula
            sheet.Range["C1"].Formula = "NewNamedRange";

            //Save and launch result file
            string result = "result.xlsx";
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

        private void btnClose_Click(object sender, EventArgs e)
        {
            Close();
        }
    }


}
