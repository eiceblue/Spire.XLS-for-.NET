using Spire.Xls;
using System;
using System.Windows.Forms;

namespace CompressPictures
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Load an existing Excel document from disk
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\CompressPictures.xlsx");

            // Get the first worksheet in the workbook
            Worksheet sheet1 = workbook.Worksheets[0];

            // Compress the picture quality for all pictures in all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                foreach (ExcelPicture picture in sheet.Pictures)
                {
                    // Set the compression level to 50 (50% of original quality)
                    picture.Compress(50);
                }
            }

            // Specify the output file name
            string result = "CompressPictures_result.xlsx";

            // Save the modified workbook to the specified file using Excel 2010 format
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
