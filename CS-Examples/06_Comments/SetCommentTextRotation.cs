using System;
using System.Data.OleDb;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.IO;
using Spire.Xls;
using Spire.Xls.Charts;
using System.Collections.Generic;
using Spire.Xls.Core.Spreadsheet.PivotTables;
using System.Drawing.Imaging;
using Spire.Xls.Core;
using Spire.Xls.Core.Spreadsheet;
using System.Text;
using Spire.Xls.Core.Spreadsheet.Collections;

namespace SetCommentTextRotation
{
	public partial class Form1 : Form
	{
        public Form1()
        {
            InitializeComponent();
        }
		private void btnRun_Click(object sender, System.EventArgs e)
        {
            //Create a workbook
            Workbook workbook = new Workbook();

            //Load the Excel document from disk
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\CellValues.xlsx");

            //Get the default first  worksheet
            Worksheet sheet = workbook.Worksheets[0];

            //Create Excel font
            ExcelFont font = workbook.CreateFont();
            font.FontName = "Arial";
            font.Size = 11;
            font.KnownColor = ExcelColors.Orange;

            //Add the comment
            CellRange range = sheet.Range["E1"];
            range.Comment.Text = "This is a comment";
            range.Comment.RichText.SetFont(0, (range.Comment.Text.Length - 1), font);

            // Set its vertical and horizontal alignment 
            range.Comment.VAlignment = CommentVAlignType.Center;
            range.Comment.HAlignment = CommentHAlignType.Right;

            //Set the comment text rotation
            range.Comment.TextRotation = TextRotationType.LeftToRight;
       
            //String for output file 
            String outputFile = "Output.xlsx";

            //Save the file
            workbook.SaveToFile(outputFile, ExcelVersion.Version2013);

            //Launching the output file.
            Viewer(outputFile);
		}
		private void Viewer( string fileName )
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
