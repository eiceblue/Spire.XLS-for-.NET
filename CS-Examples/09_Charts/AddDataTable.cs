using System;
using System.Windows.Forms;
using Spire.Xls;

namespace AddDataTable
{
    /// <summary>
    /// Summary description for Form1.
    /// </summary>
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

            // Load file from disk
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\AddDataTable.xlsx");

            //Get the first sheet
            Worksheet sheet = workbook.Worksheets[0];

            //Get the first chart
            Chart chart = sheet.Charts[0];

            // Enable the data table for the chart
            chart.HasDataTable = true;

            //Save the file 
            workbook.SaveToFile("Output.xlsx", FileFormat.Version2010);

            // Dispose of the workbook object to release resources
            workbook.Dispose();

            // Launch the file
            ExcelDocViewer("Output.xlsx");
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
