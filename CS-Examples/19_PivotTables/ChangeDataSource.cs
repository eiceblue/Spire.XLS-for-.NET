using Spire.Xls;
using System;
using System.Windows.Forms;

namespace ChangeDataSource
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
            //Load an excel file
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\ChangeDataSource.xlsx");

            //Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            CellRange Range = sheet.Range["A1:C15"];

            PivotTable table = workbook.Worksheets[1].PivotTables[0] as PivotTable;

            //Change data source
            table.ChangeDataSource(Range);
            table.Cache.IsRefreshOnLoad = false;

            string result = "ChangeDataSource_result.xlsx";
            //Save to file
            workbook.SaveToFile(result, ExcelVersion.Version2010);
            //View the document
            FileViewer(result);
         
        }

        private void FileViewer(string fileName)
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
