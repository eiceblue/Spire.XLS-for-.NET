using System;
using System.Windows.Forms;
using Spire.Xls;

namespace HideOrShowComment
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

            //Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            //Hide the second comment
            sheet.Comments[1].IsVisible = false;

            //Show the third comment
            sheet.Comments[2].IsVisible = true;


            //Save the document
            string output = "HideOrShowComment.xlsx";
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
