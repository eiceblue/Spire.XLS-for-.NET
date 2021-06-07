using Spire.Xls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace ReplaceAndHighlight
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
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\ReplaceAndHighlight.xlsx");

            Worksheet worksheet = workbook.Worksheets[0];

            CellRange[] ranges = worksheet.FindAllString("Total", true, true);

            foreach (CellRange range in ranges)
            {
                //reset the text, in other words, replace the text
                range.Text = "Sum";

                //set the color
                range.Style.Color = Color.Yellow;
            }

            string result="ReplaceAndHighlight_result.xlsx";
            workbook.SaveToFile(result, ExcelVersion.Version2010);
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
