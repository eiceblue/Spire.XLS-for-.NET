using Spire.Xls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace ToCSVWithDoubleQuotes
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            // Create a workbook
            Workbook workbook = new Workbook();

            // Load a excel document
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\ToCSV.xlsx");

            // Convert to CSV file,
            // When the last parameter is set to true, there are double quotes. The default parameter is flase
            workbook.SaveToFile("ToCSVAddQuotation.csv", ",", true);

            // Dispose of the workbook object to release resources
            workbook.Dispose();

            // Launsh the document
            ExcelDocViewer("ToCSVAddQuotation.csv");
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
