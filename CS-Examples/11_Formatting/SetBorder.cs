using System;
using System.Data.OleDb;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;

using Spire.Xls;

namespace SetBorder
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
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\SetBorder.xlsx");

            //Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            //Get the cell range where you want to apply border style
            CellRange cr = sheet.Range[sheet.FirstRow, sheet.FirstColumn, sheet.LastRow, sheet.LastColumn];

            //Apply border style 
            cr.Borders.LineStyle = LineStyleType.Double;
            cr.Borders[BordersLineType.DiagonalDown].LineStyle = LineStyleType.None;
            cr.Borders[BordersLineType.DiagonalUp].LineStyle = LineStyleType.None;
            cr.Borders.Color = Color.CadetBlue;

            string result = "SetBorder_result.xlsx";
            //Save the document
            workbook.SaveToFile(result,ExcelVersion.Version2010);

            //Launch the Excel file
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
