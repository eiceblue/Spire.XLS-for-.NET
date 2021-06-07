using Spire.Xls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace ToCSVWithFilteredValue
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
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\AutofilterSample.xlsx");
       
            //Convert to CSV file with filtered value
            workbook.Worksheets[0].SaveToFile("ToCSVWithFilteredValue.csv", ";", false);

            //Convert to CSV stream
            //worksheet.SaveToStream(Stream stream, string separator, bool retainHiddenData);           
            
            //View the document
            FileViewer("ToCSVWithFilteredValue.csv");
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
