using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using Spire.Xls;
using Spire.Xls.Core;

namespace ListDataValidation
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
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\DataValidation.xlsx");

            //Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            //Set text for cells 
            sheet.Range["A7"].Text = "Beijing";
            sheet.Range["A8"].Text = "New York";
            sheet.Range["A9"].Text = "Denver";
            sheet.Range["A10"].Text = "Paris";

            //Set data validation for cell
            CellRange range = sheet.Range["D10"];
            range.DataValidation.ShowError = true;
            range.DataValidation.AlertStyle = AlertStyleType.Stop;
            range.DataValidation.ErrorTitle = "Error";
            range.DataValidation.ErrorMessage = "Please select a city from the list";
            range.DataValidation.DataRange = sheet.Range["A7:A10"];

            //Save the document
            string output = "ListDataValidation_out.xlsx";
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
