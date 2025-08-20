using System;
using System.Data.OleDb;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;

using Spire.Xls;

namespace ReadHyperlinks
{
	public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
		private void btnRun_Click(object sender, System.EventArgs e)
		{
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Load an existing workbook from a file
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\ReadHyperlinks.xlsx");

            // Get the first worksheet in the workbook
            Worksheet sheet = workbook.Worksheets[0];

            // Retrieve the address of the first hyperlink and assign it to textBox1
            textBox1.Text = sheet.HyperLinks[0].Address;

            // Retrieve the address of the second hyperlink and assign it to textBox2
            textBox2.Text = sheet.HyperLinks[1].Address;

            // Dispose of the workbook object to release resources
            workbook.Dispose();
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
