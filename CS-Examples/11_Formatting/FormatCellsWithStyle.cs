using System;
using System.Data.OleDb;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;

using Spire.Xls;

namespace FormatCellsWithStyle
{
    public partial class Form1 : Form
    {

        public Form1()
        {
            InitializeComponent();
        }

        private void btnRun_Click(object sender, System.EventArgs e)
        {
            //Load the document from disk
            Workbook workbook = new Workbook();
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\SampleB_2.xlsx");

            //Create a style
            CellStyle style = workbook.Styles.Add("newStyle");
            //Set the shading color
            style.Color = Color.DarkGray;
            //Set the font color
            style.Font.Color = Color.White;
            //Set font name
            style.Font.FontName = "Times New Roman";
            //Set font size
            style.Font.Size = 12;
            //Set bold for the font
            style.Font.IsBold = true;
            //Set text rotation
            style.Rotation = 45;
            //Set alignment
            style.HorizontalAlignment = HorizontalAlignType.Center;
            style.VerticalAlignment = VerticalAlignType.Center;

            //Set the style for the specific range
            workbook.Worksheets[0].Range["A1:J1"].CellStyleName = style.Name;

            //Save and launch result file
            string result = "result.xlsx";
            workbook.SaveToFile(result, ExcelVersion.Version2010);
            ExcelDocViewer(result);
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
