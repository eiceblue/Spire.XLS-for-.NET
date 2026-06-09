using Spire.Xls;
using System;
using System.IO;
using System.Windows.Forms;

namespace SpecificCellsToHtml
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
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\ConversionSample1.xlsx");

            //Get the first worksheet in Excel file
            Worksheet sheet = workbook.Worksheets[0];

            // Get the specific cell range A1:B3 from the worksheet
            CellRange cell = sheet.Range["A1:E7"];

            // Extract the HTML representation of the selected cell range
            string html = cell.HtmlString;

            // Specify the output file name for the HTML result
            string result = "SpecificCellsToHtml_out.html";

            // Write the HTML content to the output file
            File.WriteAllText(result, html);

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

        private void Form1_Load(object sender, EventArgs e)
        {

        }
    }
}
