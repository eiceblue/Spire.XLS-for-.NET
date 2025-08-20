using Spire.Xls;
using System;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace MoveChartsheet
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            // Create a workbook
            Workbook workbook = new Workbook();

            // Load the document from disk
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\MoveChartsheet.xlsx");

            //Move the first chartsheet to the position of the third sheet(including chartsheet and worksheet) 
            workbook.Chartsheets[0].MoveSheet(2);

            //Move the first sheet to the position of the first chartsheet
            workbook.Chartsheets[0].MoveChartsheet(0);

            //Save to file
            string result = "MoveChartSheetResult.xlsx";
            workbook.SaveToFile(result, ExcelVersion.Version2013);

            // Dispose of the workbook object to release resources 
            workbook.Dispose();

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
