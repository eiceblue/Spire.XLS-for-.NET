using Spire.Xls;
using System;
using System.Windows.Forms;

namespace InsertHtmlStringIntoCell
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

            //Get the first sheet
            Worksheet sheet = workbook.Worksheets[0];

            // Inset html code to cell "A1"
            String htmlCode = "<div>first line<br>second line<br>third line</div>";
            CellRange range = sheet["A1"];
            range.HtmlString = htmlCode;

            // Specify the name for the resulting Excel file
            String result = "InsertHtmlStringIntoCell.xlsx";

            // Save the file
            workbook.SaveToFile(result, ExcelVersion.Version2010);

            // Dispose of the workbook object to free up resources
            workbook.Dispose();

            // Launch the document
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
