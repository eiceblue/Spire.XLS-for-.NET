using Spire.Xls;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.Windows.Forms;

namespace AllShapesToImage
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

            //Load an excel file
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\Shape.xlsx");

            //Get the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Save all shape to images
            SaveShapeTypeOption shapelist = new SaveShapeTypeOption();
            shapelist.SaveAll = true;
            List<Bitmap> images = worksheet.SaveShapesToImage(shapelist);
            int index = 0;

            // Save all images
            foreach (Image img in images)
            {
                string imageFileName = "Image_" + index + ".png";
                img.Save(imageFileName, ImageFormat.Png);
                index++;
                OutputViewer(imageFileName);
            }
            // Dispose of the workbook object to release resources
            workbook.Dispose();
        }

        private void OutputViewer(string filename)
        {
            try
            {
                System.Diagnostics.Process.Start(filename);
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
