using Spire.Xls;
using Spire.Xls.Core;
using Spire.Xls.Core.Spreadsheet.PivotTables;
using System;
using System.IO;
using System.Reflection.Emit;
using System.Windows.Forms;

namespace AddFiltersToFields
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            string outputFile = "output.xlsx";
            // Create a new workbook object
            Workbook workbook = new Workbook();

            //Load the file from disk.
             workbook.LoadFromFile(@"..\..\..\..\..\..\Data\PivotTableExample.xlsx");

            //Retrieve the first pivot table from the second sheet
            XlsPivotTable pt = workbook.Worksheets[1].PivotTables[0] as XlsPivotTable;

            //Add a label filter to the first row field of the pivot table
            pt.RowFields[0].AddLabelFilter(PivotLabelFilterType.Between, "Argentina", "Nicaragua");

            // Add a value filter on the first row field of the pivot table
            pt.RowFields[0].AddValueFilter(PivotValueFilterType.LessThan, pt.DataFields[0], 5300000, null);

             //  pt.ColumnFields[0].AddLabelFilter(PivotLabelFilterType.Between, "Argentina", "Nicaragua");
             //  pt.ColumnFields[0].AddValueFilter(PivotValueFilterType.LessThan, pt.DataFields[0], 5300000, null);

            pt.CalculateData();

            MessageBox.Show(pt.DataFields[0].Name);

            workbook.SaveToFile(outputFile, ExcelVersion.Version2013);
       
            // Dispose of the workbook object
            workbook.Dispose();

            FileViewer(outputFile);

            this.Close();
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

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }
    }
}
