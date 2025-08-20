using System;
using System.Data.OleDb;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using Spire.Xls;

namespace ReadFormulas
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
            // Load an existing workbook from a file
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\ReadFormulas.xlsx");

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Get the formula from cell C14
            string formula = sheet.Range["C14"].Formula;

            // Get the numeric value resulting from the formula in cell C14
            string formulaNumberValue = sheet.Range["C14"].FormulaNumberValue.ToString();

           // Show the formula and its numeric value
            textBox1.Text = sheet.Range["C14"].Formula;
			textBox2.Text = sheet.Range["C14"].FormulaNumberValue.ToString();

            // Dispose of the workbook object to release resources
            workbook.Dispose();
        }
		private void button1_Click(object sender, System.EventArgs e)
		{
            this.ExcelDocViewer(@"..\..\..\..\..\..\Data\ReadFormulas.xlsx");
		}

		private void ExcelDocViewer( string fileName )
		{
			try
			{
				System.Diagnostics.Process.Start(fileName);
			}
			catch{}
		}

        private void btnClose_Click(object sender, EventArgs e)
        {
            Close();
        }


	}
}
