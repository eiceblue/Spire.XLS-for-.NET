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
            //Create a workbook
            Workbook workbook = new Workbook();
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\NumberStyles.xlsx");
            //Initialize the workbook
            Worksheet sheet = workbook.Worksheets[0];

            //Input a number value for the specified cell and set the number format
            sheet.Range["B10"].Text = "NUMBER FORMATTING";
            sheet.Range["B10"].Style.Font.IsBold = true;

            sheet.Range["B13"].Text = "0";
            sheet.Range["C13"].NumberValue = 1234.5678;
            sheet.Range["C13"].NumberFormat = "0";

            sheet.Range["B14"].Text = "0.00";
            sheet.Range["C14"].NumberValue = 1234.5678;
            sheet.Range["C14"].NumberFormat = "0.00";

            sheet.Range["B15"].Text = "#,##0.00";
            sheet.Range["C15"].NumberValue = 1234.5678;
            sheet.Range["C15"].NumberFormat = "#,##0.00";

            sheet.Range["B16"].Text = "$#,##0.00";
            sheet.Range["C16"].NumberValue = 1234.5678;
            sheet.Range["C16"].NumberFormat = "$#,##0.00";

            sheet.Range["B17"].Text = "0;[Red]-0";
            sheet.Range["C17"].NumberValue = -1234.5678;
            sheet.Range["C17"].NumberFormat = "0;[Red]-0";

            sheet.Range["B18"].Text = "0.00;[Red]-0.00";
            sheet.Range["C18"].NumberValue = -1234.5678;
            sheet.Range["C18"].NumberFormat = "0.00;[Red]-0.00";

            sheet.Range["B19"].Text = "#,##0;[Red]-#,##0";
            sheet.Range["C19"].NumberValue = -1234.5678;
            sheet.Range["C19"].NumberFormat = "#,##0;[Red]-#,##0";

            sheet.Range["B20"].Text = "#,##0.00;[Red]-#,##0.000";
            sheet.Range["C20"].NumberValue = -1234.5678;
            sheet.Range["C20"].NumberFormat = "#,##0.00;[Red]-#,##0.00";

            sheet.Range["B21"].Text = "0.00E+00";
            sheet.Range["C21"].NumberValue = 1234.5678;
            sheet.Range["C21"].NumberFormat = "0.00E+00";

            sheet.Range["B22"].Text = "0.00%";
            sheet.Range["C22"].NumberValue = 1234.5678;
            sheet.Range["C22"].NumberFormat = "0.00%";

            sheet.Range["B13:B22"].Style.KnownColor = ExcelColors.Gray25Percent;

            //AutoFit Column
            sheet.AutoFitColumn(2);
            sheet.AutoFitColumn(3);

            String result = "Result-NumberStyles.xlsx";

            //Save and Launch
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
