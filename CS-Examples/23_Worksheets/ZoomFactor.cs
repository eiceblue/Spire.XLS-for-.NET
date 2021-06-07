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
			//Create a workbook and load a file
            Workbook workbook = new Workbook();
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\ZoomFactor.xlsx");
			
			//Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
			
            //Set the zoom factor of the sheet to 85
            sheet.Zoom = 85;
			
			//Save the document and launch it
            workbook.SaveToFile("ZoomFactor_result.xlsx",ExcelVersion.Version2010);
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
