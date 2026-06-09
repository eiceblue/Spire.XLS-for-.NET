using System;
using System.Data.OleDb;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;

using Spire.Xls;

namespace PDURATIONFunction
{
	public partial class Form1 : Form
	{
        public Form1()
        {
            InitializeComponent();
        }
		private void btnRun_Click(object sender, System.EventArgs e)
		{

            //Create workbook instance
            Workbook workbook = new Workbook();

            //Get first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            //Set headers
            sheet.Range["A1"].Text = "Financial Calculation";
            sheet.Range["A2"].Text = "Annual Interest Rate";
            sheet.Range["A3"].Text = "Present Value (PV)";
            sheet.Range["A4"].Text = "Future Value (FV)";
            sheet.Range["A5"].Text = "Required Periods (Years)";

            //Set input values
            sheet.Range["B2"].Text = "2.5%";
            sheet.Range["B3"].Text = "2000";
            sheet.Range["B4"].Text = "2200";

            //Calculate years to grow from 2000 to 2200 at 2.5% annual rate using PDURATION function
            sheet.Range["B5"].Formula = "=PDURATION(2.5%,2000,2200)";  
            //Format cells
            sheet.Range["B2"].Style.NumberFormat = "0.00%";
            sheet.Range["B3:B4"].Style.NumberFormat = "#,##0";
            sheet.Range["B5"].Style.NumberFormat = "0.00";

            //Auto fit columns
            sheet.AllocatedRange.AutoFitColumns();

            // Specify the file name for the resulting Excel file
            String result = "PDURATION_Calculation.xlsx";

            // Save the workbook to the specified file in Excel 2013 format
            workbook.SaveToFile(result, ExcelVersion.Version2013);

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

        private void Form1_Load(object sender, EventArgs e)
        {

        }
    }
}
