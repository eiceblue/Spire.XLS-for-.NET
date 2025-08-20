using System;
using System.Data.OleDb;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.IO;
using System.Windows.Forms;
using System.Data;

using Spire.Xls;
using Spire.Xls.Charts;

namespace OfficeOpenXMLToExcel
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

            // Load file from xml
            using (FileStream fileStream = File.OpenRead(@"..\..\..\..\..\..\Data\OfficeOpenXMLToExcel.Xml"))
            {
                workbook.LoadFromXml(fileStream);
            }

            // Save to Excel file
            workbook.SaveToFile("OfficeOpenXMLToExcel.xlsx", ExcelVersion.Version2010);

            // Dispose of the workbook object to release resources
            workbook.Dispose();

            // Launch the file 
            ExcelDocViewer("OfficeOpenXMLToExcel.xlsx");
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
