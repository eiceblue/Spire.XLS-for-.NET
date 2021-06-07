using System;
using System.Data.OleDb;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;

using Spire.Xls;

namespace XLSB
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        private void btnLoad_Click(object sender, System.EventArgs e)
        {
            Workbook workbook = new Workbook();

            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\XLSB.xlsb");
            //Initailize worksheet
            Worksheet sheet = workbook.Worksheets[0];

            this.dataGrid1.DataSource = sheet.ExportDataTable();
            this.btnSave.Enabled = true;
        }

        private void btnSave_Click(object sender, System.EventArgs e)
        {

            Workbook workbook = new Workbook();

            //Initailize worksheet
            Worksheet sheet = workbook.Worksheets[0];

            sheet.InsertDataTable((DataTable)this.dataGrid1.DataSource, true,1, 1, -1, -1);

            //Sets body style
            CellStyle oddStyle = workbook.Styles.Add("oddStyle");
            oddStyle.Borders[BordersLineType.EdgeLeft].LineStyle = LineStyleType.Thin;
            oddStyle.Borders[BordersLineType.EdgeRight].LineStyle = LineStyleType.Thin;
            oddStyle.Borders[BordersLineType.EdgeTop].LineStyle = LineStyleType.Thin;
            oddStyle.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.Thin;
            oddStyle.KnownColor = ExcelColors.LightGreen1;

            CellStyle evenStyle = workbook.Styles.Add("evenStyle");
            evenStyle.Borders[BordersLineType.EdgeLeft].LineStyle = LineStyleType.Thin;
            evenStyle.Borders[BordersLineType.EdgeRight].LineStyle = LineStyleType.Thin;
            evenStyle.Borders[BordersLineType.EdgeTop].LineStyle = LineStyleType.Thin;
            evenStyle.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.Thin;
            evenStyle.KnownColor = ExcelColors.LightTurquoise;

            foreach (CellRange range in sheet.AllocatedRange.Rows)
            {
                if (range.Row % 2 == 0)
                    range.CellStyleName = evenStyle.Name;
                else
                    range.CellStyleName = oddStyle.Name;
            }

            //Sets header style
            CellStyle styleHeader = sheet.AllocatedRange.Rows[0].Style;
            styleHeader.Borders[BordersLineType.EdgeLeft].LineStyle = LineStyleType.Thin;
            styleHeader.Borders[BordersLineType.EdgeRight].LineStyle = LineStyleType.Thin;
            styleHeader.Borders[BordersLineType.EdgeTop].LineStyle = LineStyleType.Thin;
            styleHeader.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.Thin;
            styleHeader.VerticalAlignment = VerticalAlignType.Center;
            styleHeader.KnownColor = ExcelColors.Green;
            styleHeader.Font.KnownColor = ExcelColors.White;
            styleHeader.Font.IsBold = true;

            sheet.AllocatedRange.AutoFitColumns();
            sheet.AllocatedRange.AutoFitRows();

            sheet.Rows[0].RowHeight = 20;

            workbook.SaveToFile("sample.xlsb", ExcelVersion.Xlsb2010);
            ExcelDocViewer("sample.xlsb");
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