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
			Workbook workbook = new Workbook();
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\ReadFormulas.xlsx");
			Worksheet sheet = workbook.Worksheets[0];
            
			textBox1.Text = sheet.Range["C14"].Formula;
            textBox2.Text = sheet.Range["C14"].FormulaNumberValue.ToString();
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
