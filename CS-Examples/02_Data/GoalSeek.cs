using Spire.Xls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace GoalSeek
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            // Create a workbook
            Workbook workbook = new Workbook();

            // Get the first sheet
            Worksheet sheet = workbook.Worksheets[0];

            // Set value for cell "A1"
            sheet.Range["A1"].Value = "100";

            // Set formula for cell "A2"
            CellRange targetCell = sheet.Range["A2"];
            targetCell.Formula = "=SUM(A1+B1)";

            // Variable cell
            CellRange gussCell = sheet.Range["B1"];
            Spire.Xls.GoalSeek goalSeek = new Spire.Xls.GoalSeek();

            // Trial solution
            GoalSeekResult result = goalSeek.TryCalculate(targetCell, 500, gussCell);

            //Determine the solution
            result.Determine();

            // Save the file
            workbook.SaveToFile("GoalSeek.xlsx", ExcelVersion.Version2013);

            // Dispose of the workbook object to free up resources
            workbook.Dispose();

            // Launch the document
            ExcelDocViewer("GoalSeek.xlsx");
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
