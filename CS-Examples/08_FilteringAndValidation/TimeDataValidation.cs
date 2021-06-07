using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using Spire.Xls;
using Spire.Xls.Core;

namespace TimeDataValidation
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

            sheet.Range["C12"].Text = "Please enter time between 09:00 and 18:00:";
            sheet.Range["C12"].AutoFitColumns();

            //Set Time data validation for cell "D12"
            CellRange range = sheet.Range["D12"];
            range.DataValidation.AllowType = CellDataType.Time;
            range.DataValidation.CompareOperator = ValidationComparisonOperator.Between;

            range.DataValidation.Formula1 = "09:00";
            range.DataValidation.Formula2 = "18:00";

            range.DataValidation.AlertStyle = AlertStyleType.Info;
            range.DataValidation.ShowError = true;
            range.DataValidation.ErrorTitle = "Time Error";
            range.DataValidation.ErrorMessage = "Please enter a valid time";
            range.DataValidation.InputMessage = "Time Validation Type";
            range.DataValidation.IgnoreBlank = true;
            range.DataValidation.ShowInput = true;
            
            //Save the document
            string output = "TimeDataValidation_out.xlsx";
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
