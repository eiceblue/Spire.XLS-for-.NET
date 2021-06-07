using System;
using System.Windows.Forms;
using Spire.Xls;

namespace AddCommentWithAuthor
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
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\WorksheetSample1.xlsx");

            //Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            //Get the range that will add comment
            CellRange range = sheet.Range["C1"];

            //Set the author and comment content
            string author = "E-iceblue";
            string text = "This is demo to show how to add a comment with editable Author property.";

            //Add comment to the range and set properties
            ExcelComment comment = range.AddComment();
            comment.Width = 200;
            comment.Visible = true;
            comment.Text = string.IsNullOrEmpty(author) ? text : author + ":\n" + text;

            //Set the font of the author
            ExcelFont font = range.Worksheet.Workbook.CreateFont();
            font.FontName = "Tahoma";
            font.KnownColor = ExcelColors.Black;
            font.IsBold = true;
            comment.RichText.SetFont(0, author.Length, font);

            //Save the document
            string output = "AddCommentWithAuthor.xlsx";
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
