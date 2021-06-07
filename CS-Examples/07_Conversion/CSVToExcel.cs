using Spire.Xls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace CSVToExcel
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            //Create a workbook
            Workbook workbook = new Workbook();

            //Load a csv file
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\CSVToExcel.csv", ",", 1, 1);

            Worksheet sheet = workbook.Worksheets[0];
            sheet.Range["D2:E19"].IgnoreErrorOptions = IgnoreErrorType.NumberAsText;
            sheet.AllocatedRange.AutoFitColumns();

            //Save the document and launch it
            workbook.SaveToFile("CSVToExcel_result.xlsx", ExcelVersion.Version2013);
            ExcelDocViewer("CSVToExcel_result.xlsx");
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
