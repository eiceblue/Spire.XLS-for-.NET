using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;

using Spire.Xls;
using Spire.Xls.Core;

namespace AddScrollBarControl
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
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\ExcelSample_N1.xlsx");

            //Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            //Set a value for range B10
            sheet.Range["B10"].Value2 = 1;
            sheet.Range["B10"].Style.Font.IsBold = true;

            //Add scroll bar control
            IScrollBarShape scrollBar = sheet.ScrollBarShapes.AddScrollBar(10, 3, 150, 20);
            scrollBar.LinkedCell = sheet.Range["B10"];
            scrollBar.Min = 1;
            scrollBar.Max = 150;
            scrollBar.IncrementalChange = 1;
            scrollBar.Display3DShading = true;

            //Save the document
            string output = "AddScrollBarControl_out.xlsx";
            workbook.SaveToFile(output, ExcelVersion.Version2013);

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
