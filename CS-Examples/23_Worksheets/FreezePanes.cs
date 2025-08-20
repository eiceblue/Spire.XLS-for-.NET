using System;
using System.Data.OleDb;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;

using Spire.Xls;
using Spire.Xls.Charts;

namespace FreezePanes
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

			// Load file from disk
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\FreezePanes.xlsx");

            // Get the first sheet
			Worksheet sheet = workbook.Worksheets[0];

            // Freeze Top Row
			sheet.FreezePanes(2,1);

            // Save the modified target workbook to a file
            workbook.SaveToFile("Output.xlsx",ExcelVersion.Version2010);

            // Dispose of the workbook object to release resources
            workbook.Dispose();

            // Launch the file
            ExcelDocViewer("Output.xlsx");
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
