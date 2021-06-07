using System;
using System.Data.OleDb;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;

using Spire.Xls;

namespace ReadComment
{
	public partial class Form1 : Form
	{
        public Form1()
        {
            InitializeComponent();
        }
		private void btnRun_Click(object sender, System.EventArgs e)
		{
			Workbook workbook = new Workbook();

                      workbook.LoadFromFile(@"..\..\..\..\..\..\Data\ReadComment.xls");
			
			Worksheet sheet = workbook.Worksheets[0];

			textBox1.Text = sheet.Range["A1"].Comment.Text;
			richTextBox1.Rtf = sheet.Range["A2"].Comment.RichText.RtfText;
		}

        private void btnClose_Click(object sender, EventArgs e)
        {
            Close();
        }
	}
}
