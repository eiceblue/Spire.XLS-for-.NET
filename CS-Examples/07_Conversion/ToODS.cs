using Spire.Xls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace ToODS
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

            //load a excel document
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\ToODS.xlsx");
   
            //convert to ODS file
            workbook.SaveToFile("Result.ods", FileFormat.ODS);

            //view the document
            ExcelDocViewer("Result.ods");
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
