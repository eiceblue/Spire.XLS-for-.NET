using Spire.Xls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Xls.Core.Spreadsheet;

namespace GetStyleSetStyle
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

            //Load a excel document
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\templateAz.xlsx");
            
            //Get the first sheet
            Worksheet sheet = workbook.Worksheets[0];

            //Get "B4" cell
            CellRange range = sheet.Range["B4"];       
            //Get the style of cell
            CellStyle style = range.Style;
            style.Font.FontName = "Calibri";
            style.Font.IsBold = true;
            style.Font.Size = 15;
            style.Font.Color = Color.CornflowerBlue;

            range.Style = style;

            String result = "UseGetStyleSetStyle_result.xlsx";

            //Save to file
            workbook.SaveToFile(result, ExcelVersion.Version2010);
           
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
