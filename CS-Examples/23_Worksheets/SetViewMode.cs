using System;
using System.Data.OleDb;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;

using Spire.Xls;

namespace SetViewMode
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

            // Load the document from disk
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\SetViewMode.xlsx");

            // Set the view mode as Preview
            workbook.Worksheets[0].ViewMode=ViewMode.Preview;
			
			// Save the document 
            workbook.SaveToFile("SetViewMode_result.xlsx",ExcelVersion.Version2010);

            // Dispose of the workbook object to release resources
            workbook.Dispose();

            // Launch the file
            ExcelDocViewer("SetViewMode_result.xlsx");
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
