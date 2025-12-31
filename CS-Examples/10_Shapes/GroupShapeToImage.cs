using Spire.Xls;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.Windows.Forms;

namespace GroupShapeToImage
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

            // Load an excel file
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\GroupShapeToImage.xlsx");

            // Get the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Save to image
            SaveShapeTypeOption saveShapeTypeOption = new SaveShapeTypeOption();
            saveShapeTypeOption.SaveGroupShape = true;
            List<Bitmap> images = worksheet.SaveShapesToImage(saveShapeTypeOption);
            for (int i = 0; i < images.Count; i++)
            {
                String imageFile = string.Format("Image-{0}.png", i);
                images[i].Save(imageFile, ImageFormat.Png);
            }
            
            //////////////////Use the following code for netstandard dlls/////////////////////////
            /*
            List<SkiaSharp.SKBitmap> images = worksheet.SaveShapesToImage(saveShapeTypeOption);
            for (int i = 0; i < images.Count; i++)
            { 
                SkiaSharp.SKImage image = SkiaSharp.SKImage.FromBitmap(images[i]);
                String imageFile = string.Format("Image-{0}.png", i);
                FileStream fileStream = new FileStream(imageFile, FileMode.Create, FileAccess.Write);
                image.Encode(SkiaSharp.SKEncodedImageFormat.Png, 100).SaveTo(fileStream);
            }
            */
            
            workbook.Dispose();
        }
        private void btnClose_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void FileViewer(string fileName)
        {
            try
            {
                System.Diagnostics.Process.Start(fileName);
            }
            catch { }
        }
    }
}
