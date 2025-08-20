using Spire.Xls;
using System;
using System.Windows.Forms;

namespace RemovePictureBorder
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

            //Load the Excel document from disk
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\PictureBorder.xlsx");

            // Get the first worksheet
            Worksheet sheet1 = workbook.Worksheets[0];

            // Get the first picture from the first worksheet
            ExcelPicture picture = sheet1.Pictures[0];

            // Remove the picture border
            //Method-1:
            picture.Line.Visible = false;

            // Method-2:
            //picture.Line.Weight = 0;

            // Specify the resulting file name.
            String result = "RemovePictureBorder.xlsx";

            // Save the modified workbook to a file using Excel 2013 format.
            workbook.SaveToFile(result, ExcelVersion.Version2010);

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
