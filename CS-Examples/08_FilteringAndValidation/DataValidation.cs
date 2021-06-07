using System;
using System.Data.OleDb;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;

using Spire.Xls;

namespace DataValidation
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
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\DataValidation.xlsx");
			Worksheet sheet = workbook.Worksheets[0];

            //Decimal DataValidation
			sheet.Range["B11"].Text = "Input Number(3-6):";
            CellRange rangeNumber = sheet.Range["B12"];
            //Set the operator for the data validation.
            rangeNumber.DataValidation.CompareOperator = ValidationComparisonOperator.Between;
            //Set the value or expression associated with the data validation.
            rangeNumber.DataValidation.Formula1 = "3";
            //The value or expression associated with the second part of the data validation.
            rangeNumber.DataValidation.Formula2 = "6";
            //Set the data validation type.
            rangeNumber.DataValidation.AllowType = CellDataType.Decimal;
            //Set the data validation error message.
            rangeNumber.DataValidation.ErrorMessage = "Please input correct number!";
            //Enable the error.
            rangeNumber.DataValidation.ShowError = true;
			rangeNumber.Style.KnownColor = ExcelColors.Gray25Percent;

            //Date DataValidation
            sheet.Range["B14"].Text = "Input Date:";
			CellRange rangeDate = sheet.Range["B15"];
			rangeDate.DataValidation.AllowType = CellDataType.Date;
            rangeDate.DataValidation.CompareOperator = ValidationComparisonOperator.Between;
            rangeDate.DataValidation.Formula1= "1/1/1970";
            rangeDate.DataValidation.Formula2 = "12/31/1970";
            rangeDate.DataValidation.ErrorMessage = "Please input correct date!";
			rangeDate.DataValidation.ShowError = true;
            rangeDate.DataValidation.AlertStyle = AlertStyleType.Warning;
            rangeDate.Style.KnownColor = ExcelColors.Gray25Percent;

            //TextLength DataValidation
            sheet.Range["B17"].Text = "Input Text:";
            CellRange rangeTextLength = sheet.Range["B18"];
            rangeTextLength.DataValidation.AllowType = CellDataType.TextLength;
            rangeTextLength.DataValidation.CompareOperator = ValidationComparisonOperator.LessOrEqual;
            rangeTextLength.DataValidation.Formula1 = "5";
            rangeTextLength.DataValidation.ErrorMessage = "Enter a Valid String!";
            rangeTextLength.DataValidation.ShowError = true;
            rangeTextLength.DataValidation.AlertStyle = AlertStyleType.Stop;
            rangeTextLength.Style.KnownColor = ExcelColors.Gray25Percent;

            sheet.AutoFitColumn(2);

            string result="DataValidation_result.xlsx";
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
