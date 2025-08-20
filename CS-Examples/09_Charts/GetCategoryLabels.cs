using System;
using System.Windows.Forms;
using Spire.Xls;
using System.Text;
using System.IO;

namespace GetCategoryLabels
{
    public partial class Form1 : Form
    {

        public Form1()
        {
            InitializeComponent();
        }

        private void btnRun_Click(object sender, System.EventArgs e)
        {
            StringBuilder sb = new StringBuilder();

            //Create a workbook
            Workbook workbook = new Workbook();

            // Load file from the disk 
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\SampeB_4.xlsx");

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Get the first chart
            Chart chart = sheet.Charts[0];

            //Get the cell range of the category labels
            CellRange cr = chart.PrimaryCategoryAxis.CategoryLabels;
            foreach (var cell in cr)
            {
                sb.Append(cell.Value + "\r\n");
            }

            //Save the result file
            string result = "result.txt";
            File.WriteAllText(result, sb.ToString());

            // Dispose of the workbook object to release resources
            workbook.Dispose();

            // Launch the file
            ExcelDocViewer(result);

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
