using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using Spire.Xls;
using Spire.Xls.Core;

namespace WholeNumberDataValidation
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

            sheet.Range["C12"].Text = "Please enter number between 10 and 100:";
            sheet.Range["C12"].AutoFitColumns();

            //Set Whole Number data validation for cell "D12"
            CellRange range = sheet.Range["D12"];
            range.DataValidation.AllowType = CellDataType.Integer;
            range.DataValidation.CompareOperator = ValidationComparisonOperator.Between;

            range.DataValidation.Formula1 = "10";
            range.DataValidation.Formula2 = "100";

            range.DataValidation.AlertStyle = AlertStyleType.Info;
            range.DataValidation.ShowError = true;
            range.DataValidation.ErrorTitle = "Error";
            range.DataValidation.ErrorMessage = "Please enter a valid number";
            range.DataValidation.InputMessage = "Whole Number Validation Type";
            range.DataValidation.IgnoreBlank = true;
            range.DataValidation.ShowInput = true;
            
            //Save the document
            string output = "WholeNumberDataValidation_out.xlsx";
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
