using Spire.Xls;
using Spire.Xls.Core.Spreadsheet.Shapes;
using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.Windows.Forms;

namespace ShapeToImage
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

            //Load the Excel document from disk
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\ShapeToImage.xlsx");

            //Get the first worksheet
            Worksheet sheet1 = workbook.Worksheets[0];

            //Get the first shape from the first worksheet
            XlsShape shape = sheet1.PrstGeomShapes[0] as XlsShape;

            //Save the shape to a image
            Image img = shape.SaveToImage();
            img.Save("ShapeToImage.png", ImageFormat.Png);


            // Dispose of the workbook object to release resources
            workbook.Dispose();

            // Launch the file
            FileViewer("ShapeToImage.png");
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
