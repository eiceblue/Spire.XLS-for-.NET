using System;
using System.Windows.Forms;
using Spire.Xls;

namespace WriteComment
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
			
			// Load file from disk
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\WriteComment.xlsx");

            // Get the default first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Creates a font
            ExcelFont font = workbook.CreateFont();
            font.FontName = "Arial";
            font.Size = 11;
            font.KnownColor = ExcelColors.Orange;
            ExcelFont fontBlue = workbook.CreateFont();
            fontBlue.KnownColor = ExcelColors.LightBlue;
            ExcelFont fontGreen = workbook.CreateFont();
            fontGreen.KnownColor = ExcelColors.LightGreen;

            // Get the cell B11
            CellRange range = sheet.Range["B11"];

            // Set text for the range
            range.Text = "Regular comment";

            // Add a regular comment to the cell
            range.Comment.Text = "Regular comment";

            // Auto fit column width for the range
            range.AutoFitColumns();

            // Get the cell B12
            range = sheet.Range["B12"];

            // Set text for the range
            range.Text = "Rich text comment";

            // Set font for the rich text in the comment
            range.RichText.SetFont(0, 16, font);

            // Auto fit column width for the range
            range.AutoFitColumns();

            // Set rich text comment for the cell
            range.Comment.RichText.Text = "Rich text comment";
            range.Comment.RichText.SetFont(0, 4, fontGreen);
            range.Comment.RichText.SetFont(5, 9, fontBlue);

            // Save the file
            string result = "WriteComment_result.xlsx";
            workbook.SaveToFile(result, ExcelVersion.Version2007);

            // Dispose of the workbook object to release resources
            workbook.Dispose();

            // Launch the file
            ExcelDocViewer(result);
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
