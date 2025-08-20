using Spire.Xls;
using Spire.Xls.Core;
using Spire.Xls.Core.Spreadsheet.PivotTables;
using System;
using System.Data;
using System.IO;
using System.Reflection.Emit;
using System.Windows.Forms;

namespace GetNamedRangeOfCellRange
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnRun_Click(object sender, EventArgs e)
        {

            string outputFile = "GetNamedRangeOfCellRange.txt";

            // Create a new workbook object
            Workbook workbook = new Workbook();

            //Load the file from disk.
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\AllNamedRanges.xlsx");
      
            // Determine whether NamedRange exists in Range A7:D7
            var result = workbook.Worksheets[0].Range["A7:D7"].GetNamedRange();
            File.WriteAllText(outputFile, "A7:D7---"+ result.Name + "\r\n");

            // Determine whether NamedRange exists in Range A4:D4
            var result1 = workbook.Worksheets[0].Range["A4:D4"].GetNamedRange();
            File.AppendAllText(outputFile, "A4:D4---"+result1.Name + "\r\n");

            // Determine whether NamedRange exists in cell C14
            var result2 = workbook.Worksheets[0].Range["C14"].GetNamedRange();
            if (result2 == null)
            {
                File.AppendAllText(outputFile, "C14 cell does not have NameRange");
            }

            workbook.CalculateAllValue();

            // Dispose of the workbook object to release resources
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
