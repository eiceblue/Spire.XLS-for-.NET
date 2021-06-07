using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;

using Spire.Xls;

namespace AutofilterNonBlank
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

            //Load the document from disk
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\AutofilterBlank.xlsx");

            //Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            //Match the non blank data
            sheet.AutoFilters.MatchNonBlanks(0);

            //Filter
            sheet.AutoFilters.Filter();

            //Save the document
            string output = "AutofilterNonBlank_out.xlsx";
            workbook.SaveToFile(output,ExcelVersion.Version2013);

            //Launch the Excel file
            ExcelDocViewer(output);
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
