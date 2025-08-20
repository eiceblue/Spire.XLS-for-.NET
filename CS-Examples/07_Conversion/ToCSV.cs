using Spire.Xls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace ToCSV
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
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\ToCSV.xlsx");

            //get the first sheet
            Worksheet sheet = workbook.Worksheets[0];

            //convert to CSV file
            sheet.SaveToFile("ToCSV.csv", ",", Encoding.UTF8);

            // Dispose of the workbook object to release resources
            workbook.Dispose();

            //view the document
            ExcelDocViewer("ToCSV.csv");
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
