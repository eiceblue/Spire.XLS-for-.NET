using System;
using System.Data.OleDb;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;

using Spire.Xls;

namespace SetDefaultRowAndColumnStyle
{
    public partial class Form1 : Form
    {

        public Form1()
        {
            InitializeComponent();
        }

        private void btnRun_Click(object sender, System.EventArgs e)
        {

            Workbook workbook = new Workbook();

            //Get the first sheet
            Worksheet sheet = workbook.Worksheets[0];

            //Create a cell style and set the color
            CellStyle style = workbook.Styles.Add("Mystyle");
            style.Color = Color.Yellow;

            //Set the default style for the first row and column 
            sheet.SetDefaultRowStyle(1, style);
            sheet.SetDefaultColumnStyle(1, style);

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
