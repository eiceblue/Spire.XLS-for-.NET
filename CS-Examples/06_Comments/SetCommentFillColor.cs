using Spire.Xls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace SetCommentFillColor
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            Workbook workbook = new Workbook();

            //Get the default first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            //Create Excel font
            ExcelFont font = workbook.CreateFont();
            font.FontName = "Arial";
            font.Size = 11;
            font.KnownColor = ExcelColors.Orange;

            //Add the comment
            CellRange range = sheet.Range["A1"];
            range.Comment.Text = "This is a comment";
            range.Comment.RichText.SetFont(0, (range.Comment.Text.Length - 1), font);

            //Set comment Color
            range.Comment.Fill.FillType = ShapeFillType.SolidColor;
            range.Comment.Fill.ForeColor = Color.SkyBlue;

            range.Comment.Visible = true;

            //String for output file 
            String result = "SetCommentFillColor_out.xlsx";

            //Save the file
            workbook.SaveToFile(result, ExcelVersion.Version2013);

            //View the document
            FileViewer(result);
        }

        private void FileViewer(string fileName)
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
