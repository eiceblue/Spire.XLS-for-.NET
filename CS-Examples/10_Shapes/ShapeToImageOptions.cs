using Spire.Xls;
using Spire.Xls.Core;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Security.Cryptography;
using System.Windows.Forms;

namespace ShapeToImageOptions
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
         
            // Create a new Workbook object
            Workbook workbook = new Workbook();

            // Load the workbook from the specified file path
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\Shape.xlsx");
            
            // Get the first worksheet from the workbook
            Worksheet sheet = workbook.Worksheets[0];

            // Convert shapes to images
            SaveShapeTypeOption shapelist = new SaveShapeTypeOption();

            // Set the option to save all shapes in the worksheet to images
            shapelist.SaveAll = true;

            // Save the shapes in the worksheet as images and store them in a dictionary
            Dictionary<IShape, Bitmap> images = sheet.SaveAndGetShapesToImage(shapelist);

            // Iterate over each shape-image pair in the dictionary
            foreach (KeyValuePair<IShape, Bitmap> pair in images)
            {
                // Get the shape and image from the pair
                IShape shape = pair.Key;
                Bitmap bitmap = pair.Value;

                // Generate a unique image file name based on shape properties
                string imageFileName = shape.Name + "_" + shape.Height + "_" + shape.Width + "_" + shape.ShapeType + ".png";

                // Save the bitmap as an image file with the generated name
                bitmap.Save(imageFileName);
           
                OutputViewer(imageFileName);
            }

            // Close Workbook
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
