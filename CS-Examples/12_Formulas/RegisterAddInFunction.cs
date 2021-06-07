using Spire.Xls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace RegisterAddInFunction
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            String input = @"..\..\..\..\..\..\Data\Test.xlam";

            //Create a workbook
            Workbook workbook = new Workbook();
            //Register AddIn function
            workbook.AddInFunctions.Add(input, "TEST_UDF");
            workbook.AddInFunctions.Add(input, "TEST_UDF1");
            //Get the first sheet
            Worksheet sheet = workbook.Worksheets[0];

            //Call AddIn function
            sheet.Range["A1"].Formula = "=TEST_UDF()";
            sheet.Range["A2"].Formula = "=TEST_UDF1()";

            String result = "RegisterAddInFunction_result.xlsx";

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
