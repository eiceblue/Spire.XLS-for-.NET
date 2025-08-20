using System;
using System.Data.OleDb;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;

using Spire.Xls;

namespace DataExport
{
	public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
		private void btnRun_Click(object sender, System.EventArgs e)
		{
            // Create a new workbook object to work with Excel files
            Workbook workbook = new Workbook();

            // Load file
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\DataExport.xlsx");

            // Get first sheet
            Worksheet sheet = workbook.Worksheets[0];

            // Export data 
            this.dataGrid1.DataSource = sheet.ExportDataTable();

            // Dispose of the workbook object to free up resources
            workbook.Dispose();
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            Close();
        }



	}
}
