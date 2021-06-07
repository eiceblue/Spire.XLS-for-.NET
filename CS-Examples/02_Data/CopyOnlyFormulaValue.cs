using Spire.Xls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace CopyOnlyFormulaValue
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            Workbook workbook = new Workbook();
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\CopyOnlyFormulaValue.xlsx");
          
            Worksheet sheet = workbook.Worksheets[0];

            //Set the copy option
            CopyRangeOptions copyOptions = CopyRangeOptions.OnlyCopyFormulaValue;

            //Copy range
            sheet.Copy(sheet.Range["A2:C2"], sheet.Range["A5:C5"], copyOptions);

            //Save to file
            workbook.SaveToFile("result.xlsx", ExcelVersion.Version2010);

            //View the document
            FileViewer("result.xlsx");
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
