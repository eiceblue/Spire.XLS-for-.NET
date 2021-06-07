using Spire.Xls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Xls.Core.Spreadsheet;

namespace UsePredefinedStyles
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
        
            //Get the first sheet
            Worksheet sheet = workbook.Worksheets[0];

            //Create a new style
            CellStyle style = workbook.Styles.Add("newStyle");
            style.Font.FontName = "Calibri";
            style.Font.IsBold = true;
            style.Font.Size = 15;
            style.Font.Color = Color.CornflowerBlue;

            //Get "B5" cell
            CellRange range =sheet.Range["B5"];          
            range.Text = "Welcome to use Spire.XLS";
            range.CellStyleName = style.Name;
            range.AutoFitColumns();

            String result = "UsePredefinedStyles_result.xlsx";

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
