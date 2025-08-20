using System;
using System.Windows.Forms;
using Spire.Xls;


namespace HtmlToExcel
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        private void btnRun_Click(object sender, System.EventArgs e)
        {
            //Html path
            string filePath = @"..\..\..\..\..\..\Data\HtmlToExcel.html";

            //Create a workbook
            Workbook workbook = new Workbook();

            //Load html
            workbook.LoadFromHtml(filePath);

            //Save to Excel file
            string result = "HtmlToExcel_result.xlsx";
            workbook.SaveToFile(result, ExcelVersion.Version2013);

            // Dispose of the workbook object to release resources
            workbook.Dispose();

            //Launch the file
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
