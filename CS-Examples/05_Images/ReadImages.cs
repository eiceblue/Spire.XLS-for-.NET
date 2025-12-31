using System;
using System.Data.OleDb;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;

using Spire.Xls;

namespace ReadImages
{

    public partial class Form1 : Form
	{
        public Form1()
        {
            InitializeComponent();
        }
		private void btnRun_Click(object sender, System.EventArgs e)
		{
            //Create a Workbook
			Workbook workbook = new Workbook();

            //Load the document from disk
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\ReadImages.xlsx");

            //Get the first sheet
			Worksheet sheet = workbook.Worksheets[0];

            //Get the first image
			ExcelPicture pic = sheet.Pictures[0];
          
            // Show Picture in the PictureBox
            using ( Form frm1 = new Form())
			{
				PictureBox pic1 = new PictureBox();
				pic1.Image = pic.Picture;
				frm1.Width = pic.Picture.Width;
				frm1.Height = pic.Picture.Height;
				frm1.StartPosition = FormStartPosition.CenterParent;
				pic1.Dock = DockStyle.Fill;
				frm1.Controls.Add(pic1);
				frm1.ShowDialog();
			}
			
			//////////////////Use the following code for netstandard dlls/////////////////////////
            /*
            SkiaSharp.SKImage image = SkiaSharp.SKImage.FromBitmap(pic.Picture);
            FileStream fileStream = new FileStream(outputFile, FileMode.Create, FileAccess.Write);
            image.Encode(SkiaSharp.SKEncodedImageFormat.Jpeg, 100).SaveTo(fileStream);
            fileStream.Flush();
            */
            
            // Dispose of the workbook object to release resources
            workbook.Dispose();
        }
		
        private void btnClose_Click(object sender, EventArgs e)
        {
            Close();
        }


	}
}
