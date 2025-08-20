using System;
using System.IO;
using System.Windows.Forms;

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
            // Create a workbook
            Workbook workbook = new Workbook();
           
            //Load the document from disk
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\ReadComment.xls");

            //Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            //Show comment in the TextBox
            textBox1.Text = sheet.Range["A1"].Comment.Text;
            richTextBox1.Rtf = sheet.Range["A2"].Comment.RichText.RtfText;

            //Dispose of the workbook object to release resources
            workbook.Dispose();
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            Close();
        }
	}
}
