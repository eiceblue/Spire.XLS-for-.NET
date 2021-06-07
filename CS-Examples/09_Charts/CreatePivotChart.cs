using System;
using System.Windows.Forms;
using Spire.Xls;
using Spire.Xls.Core;

namespace CreatePivotChart
{

	public partial class Form1 : Form
	{
        public Form1()
        {
            InitializeComponent();
        }

		private void btnRun_Click(object sender, EventArgs e)
		{
            //Create a workbook
			Workbook workbook = new Workbook();

            //Load the document from disk
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\PivotTable.xlsx");

            //get the first worksheet
           Worksheet sheet = workbook.Worksheets[0];
           //get the first pivot table in the worksheet
           IPivotTable pivotTable = sheet.PivotTables[0];

           //create a clustered column chart based on the pivot table
           Chart chart = sheet.Charts.Add(ExcelChartType.ColumnClustered, pivotTable);
           //set chart position
           chart.TopRow = 12;
           chart.LeftColumn = 1;
           chart.RightColumn = 8;
           chart.BottomRow = 30;
           chart.ChartTitle = "Product";
           chart.PrimaryCategoryAxis.MultiLevelLable = true;

            //Save the document
            string output = "CreatePivotChart.xlsx";
	    workbook.SaveToFile(output, ExcelVersion.Version2013);

            //Launch the file
			ExcelDocViewer(output);
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
