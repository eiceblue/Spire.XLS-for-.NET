using System;
using System.Data.OleDb;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;

using Spire.Xls;

namespace WriteImages
{

	public partial class Form1 : Form
	{
        public Form1()
        {
            InitializeComponent();
        }
		private void btnRun_Click(object sender, System.EventArgs e)
		{
            //Create a Workbook
			Workbook workbook = new Workbook();
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\WriteImages.xlsx");
            //Get the first sheet
			Worksheet sheet = workbook.Worksheets[0];

            //Add an image to the specific cell
            sheet.Pictures.Add(14, 5, @"..\..\..\..\..\..\Data\SpireXls.png");

            //Save and Launch
			workbook.SaveToFile("Output.xlsx",ExcelVersion.Version2010);
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
