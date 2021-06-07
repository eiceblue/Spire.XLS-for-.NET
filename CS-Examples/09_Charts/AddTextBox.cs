using System;
using System.Data.OleDb;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;

using Spire.Xls;
using Spire.Xls.Charts;
using Spire.Xls.Core.Spreadsheet.Shapes;
using Spire.Xls.Core;

namespace AddTextBox
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        private void btnRun_Click(object sender, System.EventArgs e)
        {
            //Create a Workbook
            Workbook workbook = new Workbook();
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\AddTextBox.xlsx");
            
            //Get the first sheet
            Worksheet sheet = workbook.Worksheets[0];

            //Get the first chart
            Chart chart = sheet.Charts[0];

            //Add a Textbox
            ITextBoxLinkShape textbox = chart.Shapes.AddTextBox();
            textbox.Width = 1200;
            textbox.Height = 320;
            textbox.Left = 1000;
            textbox.Top = 480;
            textbox.Text = "This is a textbox";

            //Save and Launch
            workbook.SaveToFile("Output.xlsx", ExcelVersion.Version2010);
            ExcelDocViewer("Output.xlsx");
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
