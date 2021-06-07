using System;
using System.Data.OleDb;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;

using Spire.Xls;
using Spire.Xls.Core.Spreadsheet.ConditionalFormatting;
using Spire.Xls.Core.Spreadsheet.Collections;
using Spire.Xls.Core;

namespace SimpleConditionalFormatting
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
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\ConditionalFormatting.xlsx");
             //Get the first sheet
            Worksheet oldSheet = workbook.Worksheets[0];
            AddConditionalFormattingForExistingSheet(oldSheet);

            String result = "SimpleConditionalFormatting_result.xlsx";
            //Save and Launch
            workbook.SaveToFile(result, ExcelVersion.Version2010);
            ExcelDocViewer(result);
        }
         private void AddConditionalFormattingForExistingSheet(Worksheet sheet)
        {
            sheet.AllocatedRange.RowHeight = 15;
            sheet.AllocatedRange.ColumnWidth = 16;

            //Create conditional formatting rule
            XlsConditionalFormats xcfs1 = sheet.ConditionalFormats.Add();
            xcfs1.AddRange(sheet.Range["A1:D1"]);
            IConditionalFormat cf1 = xcfs1.AddCondition();
            cf1.FormatType = ConditionalFormatType.CellValue;
            cf1.FirstFormula = "150";
            cf1.Operator = ComparisonOperatorType.Greater;
            cf1.FontColor = Color.Red;
            cf1.BackColor = Color.LightBlue;

            XlsConditionalFormats xcfs2 = sheet.ConditionalFormats.Add();
            xcfs2.AddRange(sheet.Range["A2:D2"]);
            IConditionalFormat cf2 = xcfs2.AddCondition();
            cf2.FormatType = ConditionalFormatType.CellValue;
            cf2.FirstFormula = "300";
            cf2.Operator = ComparisonOperatorType.Less;
            //Set border color
            cf2.LeftBorderColor = Color.Pink;
            cf2.RightBorderColor = Color.Pink;
            cf2.TopBorderColor = Color.DeepSkyBlue;
            cf2.BottomBorderColor = Color.DeepSkyBlue;
            cf2.LeftBorderStyle = LineStyleType.Medium;
            cf2.RightBorderStyle = LineStyleType.Thick;
            cf2.TopBorderStyle = LineStyleType.Double;
            cf2.BottomBorderStyle = LineStyleType.Double;

            //Add data bars
            XlsConditionalFormats xcfs3 = sheet.ConditionalFormats.Add();
            xcfs3.AddRange(sheet.Range["A3:D3"]);
            IConditionalFormat cf3 = xcfs3.AddCondition();
            cf3.FormatType = ConditionalFormatType.DataBar;
            cf3.DataBar.BarColor = Color.CadetBlue;

            //Add icon sets
            XlsConditionalFormats xcfs4 = sheet.ConditionalFormats.Add();
            xcfs4.AddRange(sheet.Range["A4:D4"]);
            IConditionalFormat cf4 = xcfs4.AddCondition();
            cf4.FormatType = ConditionalFormatType.IconSet;
            cf4.IconSet.IconSetType = IconSetType.ThreeTrafficLights1;

            //Add color scales
            XlsConditionalFormats xcfs5 = sheet.ConditionalFormats.Add();
            xcfs5.AddRange(sheet.Range["A5:D5"]);
            IConditionalFormat cf5 = xcfs5.AddCondition();
            cf5.FormatType = ConditionalFormatType.ColorScale;

            //Highlight duplicate values in range "A6:D6" with BurlyWood color
            XlsConditionalFormats xcfs6 = sheet.ConditionalFormats.Add();
            xcfs6.AddRange(sheet.Range["A6:D6"]);
            IConditionalFormat cf6 = xcfs6.AddCondition();
            cf6.FormatType = ConditionalFormatType.DuplicateValues;
            cf6.BackColor = Color.BurlyWood;
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
