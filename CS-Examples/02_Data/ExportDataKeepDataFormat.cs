using Spire.Xls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace ExportDataKeepDataFormat
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

			// Load the file from disk
			workbook.LoadFromFile(@"../../../../../../Data/ExportDataKeepDataFormat.xlsx");
			
			// Get the first worksheet
			Worksheet sheet = workbook.Worksheets[0];
			
			// Export DataTable without keeping data format
			ExportTableOptions options = new ExportTableOptions();
			options.KeepDataFormat = false;
			options.RenameStrategy = RenameStrategy.Digit;

			// Export data to data table
			DataTable table = sheet.ExportDataTable(1, 1, sheet.LastDataRow, sheet.LastDataColumn, options); 
			
			// Show the data table
            this.dataGridView1.DataSource = table;
			
			// Dispose of the workbook object to free up resources
			workbook.Dispose();
        }
        private void btnClose_Click(object sender, EventArgs e)
        {
            Close();
        }
    }
}
