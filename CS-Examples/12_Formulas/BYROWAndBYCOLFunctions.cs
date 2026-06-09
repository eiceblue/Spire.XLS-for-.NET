using Spire.Xls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace BYROWAndBYCOLFunctions
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Load an existing Excel file from the specified path
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\BYROWAndBYCOLFunctions.xlsx");

            // Get the first sheet from the workbook
            Worksheet sheet = workbook.Worksheets[0];

            // Calculate average score for each row (Columns B to F) using BYROW function
            sheet.Range["G2"].Formula = "=BYROW(B2:F2, LAMBDA(row, AVERAGE(row)))";
            sheet.Range["G3"].Formula = "=BYROW(B3:F3, LAMBDA(row, AVERAGE(row)))";
            sheet.Range["G4"].Formula = "=BYROW(B4:F4, LAMBDA(row, AVERAGE(row)))";
            sheet.Range["G5"].Formula = "=BYROW(B5:F5, LAMBDA(row, AVERAGE(row)))";
            sheet.Range["G6"].Formula = "=BYROW(B6:F6, LAMBDA(row, AVERAGE(row)))";
            sheet.Range["G7"].Formula = "=BYROW(B7:F7, LAMBDA(row, AVERAGE(row)))";

            // Calculate average for each column (subject) using BYCOL function
            sheet.Range["B8"].Formula = "=BYCOL(B2:B7, LAMBDA(col, AVERAGE(col)))"; 
            sheet.Range["C8"].Formula = "=BYCOL(C2:C7, LAMBDA(col, AVERAGE(col)))"; 
            sheet.Range["D8"].Formula = "=BYCOL(D2:D7, LAMBDA(col, AVERAGE(col)))"; 
            sheet.Range["E8"].Formula = "=BYCOL(E2:E7, LAMBDA(col, AVERAGE(col)))"; 
            sheet.Range["F8"].Formula = "=BYCOL(F2:F7, LAMBDA(col, AVERAGE(col)))"; 
            sheet.Range["G8"].Formula = "=BYCOL(G2:G7, LAMBDA(col, AVERAGE(col)))"; 

            // Combined BYROW and BYCOL usage examples - Comprehensive Statistics section
            sheet.Range["I1"].Value = "Comprehensive Statistics";

            // Directly get max value
            sheet.Range["I3"].Value = "Highest Average";
            sheet.Range["J3"].Formula = "=MAX(BYROW(G2:G7, LAMBDA(row, row)))"; 

            //Directly get min value
            sheet.Range["I4"].Value = "Lowest Average";
            sheet.Range["J4"].Formula = "=MIN(BYROW(G2:G7, LAMBDA(row, row)))"; 

            // Calculate overall average across all subjects and students
            sheet.Range["I5"].Value = "Overall Subject Average";
            sheet.Range["J5"].Formula = "=BYCOL(B2:F7, LAMBDA(col, AVERAGE(col)))"; 

            // Nested scenario: BYROW nested with BYCOL
            // This formula doubles each value in the range B2:E7 and sums them by row
            sheet.Range["I7"].Formula = "=BYROW(B2:E7, LAMBDA(row, SUM(BYCOL(row, LAMBDA(col, col*2)))))";

            // Recalculate all formulas to ensure values are up to date
            sheet.CalculateAllValue();

            // Save the modified workbook to the specified file using Excel 2010 format
            string result = @"BYROWAndBYCOLFunctions_out.xlsx";
            workbook.SaveToFile(result, ExcelVersion.Version2010);

            // Dispose of the workbook object to release resources
            workbook.Dispose();

            //View the document
            FileViewer(result);

        }

        private void FileViewer(string fileName)
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
