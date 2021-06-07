using System;
using System.Data.OleDb;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;

using Spire.Xls;

namespace SetTabColor
{
	public partial class Form1 :Form
	{
        public Form1()
        {
            InitializeComponent();
        }
		private void btnRun_Click(object sender, System.EventArgs e)
		{
            //Create a workbook and load a file
            Workbook workbook = new Workbook();
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\SetTabColor.xlsx");

            //Set the tab color of first sheet to be red 
            Worksheet worksheet = workbook.Worksheets[0];
            worksheet.TabColor = Color.Red;

            //Set the tab color of first sheet to be green 
            worksheet = workbook.Worksheets[1];
            worksheet.TabColor = Color.Green;

            //Set the tab color of first sheet to be blue 
            worksheet = workbook.Worksheets[2];
            worksheet.TabColor = Color.LightBlue;

            //Save the document and launch it
            workbook.SaveToFile("SetTabColor_result.xlsx",ExcelVersion.Version2010);
            ExcelDocViewer("SetTabColor_result.xlsx");
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
