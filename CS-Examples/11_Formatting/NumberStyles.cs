using System;
using System.Data.OleDb;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;

using Spire.Xls;

namespace NumberStyles
{
    public partial class Form1 : Form
    {

        public Form1()
        {

            InitializeComponent();

        }
        private void btnRun_Click(object sender, System.EventArgs e)
        {
            // Create a workbook
            Workbook workbook = new Workbook();

            // Load file from disk
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\NumberStyles.xlsx");
            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Input a number value for the specified cell and set the number format
            sheet.Range["B10"].Text = "NUMBER FORMATTING";
            sheet.Range["B10"].Style.Font.IsBold = true;

            // Display as integer
            sheet.Range["B13"].Text = "0"; 
            sheet.Range["C13"].NumberValue = 1234.5678;
            sheet.Range["C13"].NumberFormat = "0";

            // Display as two decimal places
            sheet.Range["B14"].Text = "0.00"; 
            sheet.Range["C14"].NumberValue = 1234.5678;
            sheet.Range["C14"].NumberFormat = "0.00";

            // Display with thousand separator and two decimal places
            sheet.Range["B15"].Text = "#,##0.00"; 
            sheet.Range["C15"].NumberValue = 1234.5678;
            sheet.Range["C15"].NumberFormat = "#,##0.00";

            // Display as currency with thousand separator and two decimal places
            sheet.Range["B16"].Text = "$#,##0.00"; 
            sheet.Range["C16"].NumberValue = 1234.5678;
            sheet.Range["C16"].NumberFormat = "$#,##0.00";

            // Display positive numbers as is, negative numbers in red
            sheet.Range["B17"].Text = "0;[Red]-0"; 
            sheet.Range["C17"].NumberValue = -1234.5678;
            sheet.Range["C17"].NumberFormat = "0;[Red]-0";

            // Display positive numbers with two decimal places, negative numbers in red
            sheet.Range["B18"].Text = "0.00;[Red]-0.00"; 
            sheet.Range["C18"].NumberValue = -1234.5678;
            sheet.Range["C18"].NumberFormat = "0.00;[Red]-0.00";

            // Display positive numbers with thousand separator, negative numbers in red
            sheet.Range["B19"].Text = "#,##0;[Red]-#,##0"; 
            sheet.Range["C19"].NumberValue = -1234.5678;
            sheet.Range["C19"].NumberFormat = "#,##0;[Red]-#,##0";

            // Display positive numbers with thousand separator and two decimal places, negative numbers in red
            sheet.Range["B20"].Text = "#,##0.00;[Red]-#,##0.000";
            sheet.Range["C20"].NumberValue = -1234.5678;
            sheet.Range["C20"].NumberFormat = "#,##0.00;[Red]-#,##0.00";

            // Display as scientific notation with two decimal places
            sheet.Range["B21"].Text = "0.00E+00"; 
            sheet.Range["C21"].NumberValue = 1234.5678;
            sheet.Range["C21"].NumberFormat = "0.00E+00";

            // Display as percentage with two decimal places
            sheet.Range["B22"].Text = "0.00%"; 
            sheet.Range["C22"].NumberValue = 1234.5678;
            sheet.Range["C22"].NumberFormat = "0.00%";

            // Set background color for the range
            sheet.Range["B13:B22"].Style.KnownColor = ExcelColors.Gray25Percent; 

            // AutoFit Column
            sheet.AutoFitColumn(2); 
            sheet.AutoFitColumn(3);

            // Specify the name for the resulting Excel file
            String result = "Result-NumberStyles.xlsx";

            // Save the workbook to the specified file in Excel 2010 format
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
