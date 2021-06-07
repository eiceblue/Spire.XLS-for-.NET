using Spire.Xls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace ToPostScript
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            //create a workbook
            Workbook workbook = new Workbook();

            //load an excel document
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\ToPostScript.xlsx");

            string result = "Result.ps";
            //convert to ODS file
            workbook.SaveToFile(result, FileFormat.PostScript);

            //view the document
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
