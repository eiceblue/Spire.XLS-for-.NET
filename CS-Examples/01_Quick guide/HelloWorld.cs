using Spire.Xls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace HelloWorld
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
            sheet.Range["A1"].Text = "Hello World";

            sheet.Range["A1"].AutoFitColumns();

            String result = "HelloWorld.xlsx";

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
