using System;
using System.Data.OleDb;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;

using Spire.Xls;

namespace CutCellsToOtherPosition
{
    public partial class Form1 : Form
    {

        public Form1()
        {
            InitializeComponent();
        }

        private void btnRun_Click(object sender, System.EventArgs e)
        {
            //Load the document from disk
            Workbook workbook = new Workbook();
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\SampleB_2.xlsx");

            //Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            CellRange Ori = sheet.Range["A1:C5"];
            CellRange Dest = sheet.Range["A26:C30"];

            //Copy the range to other position
            sheet.Copy(Ori, Dest, true, true, true);

            //Remove all content in original cells
            foreach (CellRange cr in Ori)
            {
                cr.ClearAll();
            }

            //Save and launch result file
            string result = "result.xlsx";
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
