using Spire.Xls;
using System;
using System.IO;
using System.Windows.Forms;

namespace SetCropPositionForImageOfHeaderFooter
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            // Create a new instance of the Workbook class
            Workbook workbook = new Workbook();

            // Load the workbook from the specified file path
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\ImageInHeaderFooter.xlsx");

            // Get the first worksheet in the workbook
            Worksheet sheet = workbook.Worksheets[0];

            // Set the cropping values for the left header picture
            sheet.PageSetup.LeftHeaderPictureCropTop = 0.2f;
            sheet.PageSetup.LeftHeaderPictureCropBottom = 0.3f;
            sheet.PageSetup.LeftHeaderPictureCropLeft = 0.3f;
            sheet.PageSetup.LeftHeaderPictureCropRight = 0.2f;

            // Set the cropping values for the left footer picture
            sheet.PageSetup.LeftFooterPictureCropTop = 0.2f;
            sheet.PageSetup.LeftFooterPictureCropBottom = 0.3f;
            sheet.PageSetup.LeftFooterPictureCropLeft = 0.3f;
            sheet.PageSetup.LeftFooterPictureCropRight = 0.2f;

            // Set the cropping values for the center header picture
            sheet.PageSetup.CenterHeaderPictureCropTop = 0.3f;
            sheet.PageSetup.CenterHeaderPictureCropBottom = 0.4f;
            sheet.PageSetup.CenterHeaderPictureCropLeft = 0.4f;
            sheet.PageSetup.CenterHeaderPictureCropRight = 0.3f;

            // Set the cropping values for the center footer picture
            sheet.PageSetup.CenterFooterPictureCropTop = 0.3f;
            sheet.PageSetup.CenterFooterPictureCropBottom = 0.4f;
            sheet.PageSetup.CenterFooterPictureCropLeft = 0.4f;
            sheet.PageSetup.CenterFooterPictureCropRight = 0.3f;

            // Set the cropping values for the right header picture
            sheet.PageSetup.RightHeaderPictureCropTop = 0.2f;
            sheet.PageSetup.RightHeaderPictureCropBottom = 0.3f;
            sheet.PageSetup.RightHeaderPictureCropLeft = 0.9f;
            sheet.PageSetup.RightHeaderPictureCropRight = 0.4f;

            // Set the cropping values for the right footer picture
            sheet.PageSetup.RightFooterPictureCropTop = 0.2f;
            sheet.PageSetup.RightFooterPictureCropBottom = 0.3f;
            sheet.PageSetup.RightFooterPictureCropLeft = 0.9f;
            sheet.PageSetup.RightFooterPictureCropRight = 0.4f;

            // Save the workbook to the specified file path with the specified file format
            String result = @"result.xlsx";
            workbook.SaveToFile(result, FileFormat.Version2013);

            // Dispose workbook object
            workbook.Dispose();

            // Launch the file
            FileViewer(result);

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
    }
}
