using System;
using System.Data.OleDb;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;

using Spire.Xls;

namespace MarkerDesigner
{
	public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

		private void btnRun_Click(object sender, System.EventArgs e)
		{
			Workbook workbook = new Workbook();
			
			workbook.LoadFromFile(@"..\..\..\..\..\..\Data\MarkerDesigner.xls");
			DataTable dt = (DataTable)dataGrid1.DataSource;

			Worksheet sheet = workbook.Worksheets[0];
            //Fill parameter
			workbook.MarkerDesigner.AddParameter("Variable1",1234.5678);
            //Fill DataTable
			workbook.MarkerDesigner.AddDataTable("Country",dt);
			workbook.MarkerDesigner.Apply();
            //AutoFit
			sheet.AllocatedRange.AutoFitRows();
			sheet.AllocatedRange.AutoFitColumns();;


            String result = "Output_MarkerDesigner.xlsx";

            workbook.SaveToFile(result, ExcelVersion.Version2010);
            ExcelDocViewer(result);
		}

		private void ExcelDocViewer( string fileName )
		{
			try
			{
				System.Diagnostics.Process.Start(fileName);
			}
			catch{}
		}

		private void Form1_Load(object sender, System.EventArgs e)
		{
			Workbook workbook = new Workbook();

            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\MarkerDesigner-DataSample.xls");
			//Initailize worksheet
			Worksheet sheet = workbook.Worksheets[0];

            this.dataGrid1.DataSource = sheet.ExportDataTable();
		}

        private void btnClose_Click(object sender, EventArgs e)
        {
            Close();
        }


	}
}
