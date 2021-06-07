using Spire.Xls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace UseArrayR1C1Formula
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

            //Get the first sheet
            Worksheet sheet = workbook.Worksheets[0];

            sheet.Range["A1"].NumberValue = 1;
            sheet.Range["A2"].NumberValue = 2;
            sheet.Range["A3"].NumberValue = 3;
            sheet.Range["B1"].NumberValue = 4;
            sheet.Range["B2"].NumberValue = 5;
            sheet.Range["B3"].NumberValue = 6;
            sheet.Range["C1"].NumberValue = 7;
            sheet.Range["C2"].NumberValue = 8;
            sheet.Range["C3"].NumberValue = 9;

            sheet.Range["B4"].Text = "Sum:";
            sheet.Range["B4"].Style.HorizontalAlignment = HorizontalAlignType.Right;
            //Write array  R1C1 formula
            sheet.Range["C4"].FormulaArrayR1C1 = "=SUM(R[-3]C[-2]:R[-1]C)";

            //Calculate Formulas
            workbook.CalculateAllValue();

            String result = "UseArrayR1C1Formulas_result.xlsx";

            //Save to file
            workbook.SaveToFile(result, ExcelVersion.Version2010);

            //View the document
            FileViewer(result);
        }

        private void FileViewer(string fileName)
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
