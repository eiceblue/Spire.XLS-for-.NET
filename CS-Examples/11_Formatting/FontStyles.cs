using System;
using System.Data.OleDb;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using Spire.Xls;

namespace FontStyles
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
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\FontStyles.xlsx");

            //Get the first sheet
            Worksheet sheet = workbook.Worksheets[0];

            //Set font style
            sheet.Range["B1"].Style.Font.FontName = "Comic Sans MS";
            sheet.Range["B2:D2"].Style.Font.FontName = "Corbel";
            sheet.Range["B3:D7"].Style.Font.FontName = "Aleo";

            //Set font size
            sheet.Range["B1"].Style.Font.Size = 45;
            sheet.Range["B2:D3"].Style.Font.Size = 25;
            sheet.Range["B3:D7"].Style.Font.Size = 12;

            //Set excel cell data to be bold
            sheet.Range["B2:D2"].Style.Font.IsBold = true;

            //Set excel cell data to be underline
            sheet.Range["B3:B7"].Style.Font.Underline = FontUnderlineType.Single;

            //set excel cell data color
            sheet.Range["B1"].Style.Font.Color = Color.CornflowerBlue;
            sheet.Range["B2:D2"].Style.Font.Color = Color.CadetBlue;
            sheet.Range["B3:D7"].Style.Font.Color = Color.Firebrick;

            //set excel cell data to be italic
            sheet.Range["B3:D7"].Style.Font.IsItalic = true;

            //Add strikethrough
            sheet.Range["D3"].Style.Font.IsStrikethrough = true;
            sheet.Range["D7"].Style.Font.IsStrikethrough = true;

            String result = "FontStyles_output.xlsx";
            //Save and Launch
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
