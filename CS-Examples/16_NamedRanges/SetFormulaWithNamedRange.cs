using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;

using Spire.Xls;
using Spire.Xls.Core;

namespace SetFormulaWithNamedRange
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

            //Create an empty sheet
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\ExcelSample_N1.xlsx");

            //Get the sheet
            Worksheet sheet = workbook.Worksheets[0];

            //Create a named range
            INamedRange NamedRange = workbook.NameRanges.Add("MyNamedRange");
            //Refers to range
            NamedRange.RefersToRange = sheet.Range["B10:B12"];

            //Set the formula of range to named range
            sheet.Range["B13"].Formula = "=SUM(MyNamedRange)";

            //Set value of ranges
            sheet.Range["B10"].Value2=10;
            sheet.Range["B11"].Value2 = 20;
            sheet.Range["B12"].Value2 = 30;

            //Save the Excel file
            string result = "SetFormulaWithNamedRange_out.xlsx";
            workbook.SaveToFile(result, ExcelVersion.Version2013);

            //Launch the Excel file
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
