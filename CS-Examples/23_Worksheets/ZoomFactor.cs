using System;
using System.Data.OleDb;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;

using Spire.Xls;

namespace ZoomFactor
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

            // Load the Excel document from disk
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\ZoomFactor.xlsx");
			
			//Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
			
            //Set the zoom factor of the sheet to 85
            sheet.Zoom = 85;
			
			//Save the document 
            workbook.SaveToFile("ZoomFactor_result.xlsx",ExcelVersion.Version2010);

            // Dispose of the workbook object to release resources
            workbook.Dispose();

            // Launch the file
            ExcelDocViewer("ZoomFactor_result.xlsx");
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
