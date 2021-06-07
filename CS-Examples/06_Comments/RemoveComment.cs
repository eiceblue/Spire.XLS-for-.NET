using System;
using System.Windows.Forms;
using Spire.Xls;
using Spire.Xls.Collections;

namespace RemoveComment
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

            //Load the document from disk
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\CommentSample.xlsx");

            //Get all comments of the first sheet
            CommentsCollection comments = workbook.Worksheets[0].Comments;
            //Change the content of the first comment
            comments[0].Text = "This comment has been changed.";
            //Remove the second comment
            comments[1].Remove();

            //Save the document
            string output = "RemoveAndChangeComment.xlsx";
			workbook.SaveToFile(output, ExcelVersion.Version2013);

            //Launch the Excel file
			ExcelDocViewer(output);
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
