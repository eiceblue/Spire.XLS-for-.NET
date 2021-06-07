using Spire.Xls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace AddWatermark
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            //initialize a new instance of workbook and load the test file
            Workbook workbook = new Workbook();
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\AddWatermark.xlsx");

            //Insert image in a header to mimic a watermark
            Font font = new System.Drawing.Font("Arial", 40);
            String watermark = "Confidential";

            foreach (Worksheet sheet in workbook.Worksheets)
            {
                //Call DrawText() to create an image
                System.Drawing.Image imgWtrmrk = DrawText(watermark, font, System.Drawing.Color.LightCoral, System.Drawing.Color.White, sheet.PageSetup.PageHeight, sheet.PageSetup.PageWidth);

                //Set image as left header image
                sheet.PageSetup.LeftHeaderImage = imgWtrmrk;
                sheet.PageSetup.LeftHeader = "&G";

                //The watermark will only appear in this mode, it will disappear if the mode is normal
                sheet.ViewMode = ViewMode.Layout;
            }

            //Save and Launch
            workbook.SaveToFile("Output.xlsx", ExcelVersion.Version2010);
            System.Diagnostics.Process.Start("Output.xlsx");
        }

        private static System.Drawing.Image DrawText(String text, System.Drawing.Font font, Color textColor, Color backColor, double height, double width)
        {
            //Create a bitmap image with specified width and height
            Image img = new Bitmap((int)width, (int)height);
            Graphics drawing = Graphics.FromImage(img);

            //Get the size of text
            SizeF textSize = drawing.MeasureString(text, font);

            //Set rotation point
            drawing.TranslateTransform(((int)width - textSize.Width) / 2, ((int)height - textSize.Height) / 2);

            //Rotate text
            drawing.RotateTransform(-45);

            //Reset translate transform    
            drawing.TranslateTransform(-((int)width - textSize.Width) / 2, -((int)height - textSize.Height) / 2);

            //Paint the background
            drawing.Clear(backColor);

            //Create a brush for the text
            Brush textBrush = new SolidBrush(textColor);

            //Draw text on the image at center position
            drawing.DrawString(text, font, textBrush, ((int)width - textSize.Width) / 2, ((int)height - textSize.Height) / 2);
            drawing.Save();
            return img;
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
