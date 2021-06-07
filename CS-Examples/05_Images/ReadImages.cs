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
			
			using( Form frm1 = new Form())
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
		}
		private void ExcelDocViewer( string fileName )
		{
			try
			{
				System.Diagnostics.Process.Start(fileName);
			}
			catch{}
		}
        private void btnClose_Click(object sender, EventArgs e)
        {
            Close();
        }


	}
}
