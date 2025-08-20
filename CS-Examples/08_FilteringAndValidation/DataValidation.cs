using System;
using System.Windows.Forms;
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
            // Create a workbook to store Excel data
            Workbook workbook = new Workbook();

            // Load the Excel document from disk into the workbook
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\DataValidation.xlsx");

            // Get the first worksheet in the workbook
            Worksheet sheet = workbook.Worksheets[0];

            // Decimal DataValidation
            sheet.Range["B11"].Text = "Input Number(3-6):";
            CellRange rangeNumber = sheet.Range["B12"];
            rangeNumber.DataValidation.CompareOperator = ValidationComparisonOperator.Between;
            rangeNumber.DataValidation.Formula1 = "3";
            rangeNumber.DataValidation.Formula2 = "6";
            rangeNumber.DataValidation.AllowType = CellDataType.Decimal;
            rangeNumber.DataValidation.ErrorMessage = "Please input correct number!";
            rangeNumber.DataValidation.ShowError = true;
            rangeNumber.Style.KnownColor = ExcelColors.Gray25Percent;

            // Date DataValidation
            sheet.Range["B14"].Text = "Input Date:";
            CellRange rangeDate = sheet.Range["B15"];
            rangeDate.DataValidation.AllowType = CellDataType.Date;
            rangeDate.DataValidation.CompareOperator = ValidationComparisonOperator.Between;
            rangeDate.DataValidation.Formula1 = "1/1/1970";
            rangeDate.DataValidation.Formula2 = "12/31/1970";
            rangeDate.DataValidation.ErrorMessage = "Please input correct date!";
            rangeDate.DataValidation.ShowError = true;
            rangeDate.DataValidation.AlertStyle = AlertStyleType.Warning;
            rangeDate.Style.KnownColor = ExcelColors.Gray25Percent;

            // TextLength DataValidation
            sheet.Range["B17"].Text = "Input Text:";
            CellRange rangeTextLength = sheet.Range["B18"];
            rangeTextLength.DataValidation.AllowType = CellDataType.TextLength;
            rangeTextLength.DataValidation.CompareOperator = ValidationComparisonOperator.LessOrEqual;
            rangeTextLength.DataValidation.Formula1 = "5";
            rangeTextLength.DataValidation.ErrorMessage = "Enter a Valid String!";
            rangeTextLength.DataValidation.ShowError = true;
            rangeTextLength.DataValidation.AlertStyle = AlertStyleType.Stop;
            rangeTextLength.Style.KnownColor = ExcelColors.Gray25Percent;

            // Auto fit the column width for better visibility
            sheet.AutoFitColumn(2);

            // Save the modified workbook with data validations to a new file named "DataValidation_result.xlsx"
            string result = "DataValidation_result.xlsx";
            workbook.SaveToFile(result, ExcelVersion.Version2010);

            // Dispose of the workbook object to release resources
            workbook.Dispose();

            // Launch the file
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
