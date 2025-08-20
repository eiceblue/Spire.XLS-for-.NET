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
            //Create a workbook.
            Workbook workbook = new Workbook();

            //Load the document from disk
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\SampleB_2.xlsx");

            //Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Define the original source range to be copied (cells A1 to C5)
            CellRange Ori = sheet.Range["A1:C5"];

            // Define the destination range where the source range will be copied to (cells A26 to C30)
            CellRange Dest = sheet.Range["A26:C30"];

            //Copy the range to other position
            sheet.Copy(Ori, Dest, true, true, true);

            //Remove all content in original cells
            foreach (CellRange cr in Ori)
            {
                cr.ClearAll();
            }

            ///Specify the filename for the resulting Excel file
            string result = "result.xlsx";

            // Save the workbook to the specified file in Excel 2010 format
            workbook.SaveToFile(result, ExcelVersion.Version2010);

            // Dispose of the workbook object to release resources
            workbook.Dispose();

            // View file
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
