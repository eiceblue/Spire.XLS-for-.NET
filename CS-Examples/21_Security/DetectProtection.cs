using System;
using System.Data.OleDb;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;

using Spire.Xls;

namespace DetectProtection
{

	public partial class Form1 : Form
	{
        public Form1()
        {
            InitializeComponent();
        }

		private void btnRun_Click(object sender, EventArgs e)
		{
            // Specify the input file
            string input = @"..\..\..\..\..\..\Data\ProtectedWorkbook.xlsx";

            //Detect if the Excel workbook is password protected
            bool value = Workbook.IsPasswordProtected(input);

            if (value)
            {
                textBox1.Text = "Yes";
            }
            else
            {
                textBox1.Text = "No";
            }

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
