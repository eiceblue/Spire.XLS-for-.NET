using Spire.Xls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace SetHeaderFooter
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            //Load a Workbook from disk        
            Workbook Workbook = new Workbook();
            Workbook.LoadFromFile(@"..\..\..\..\..\..\Data\SetHeaderFooter.xlsx");

            //Get the first worksheet
            Worksheet Worksheet = Workbook.Worksheets[0];


            //Set left header,"Arial Unicode MS" is font name, "18" is font size.
            Worksheet.PageSetup.LeftHeader = "&\"Arial Unicode MS\"&14 Spire.XLS for .NET ";

            //Set center footer 
            Worksheet.PageSetup.CenterFooter = "Footer Text";

            Worksheet.ViewMode = ViewMode.Layout;

            String result = "SetHeaderFooter_result.xlsx";
            //Save and Launch
            Workbook.SaveToFile(result, ExcelVersion.Version2010);
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
