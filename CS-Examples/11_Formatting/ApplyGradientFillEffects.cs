using Spire.Xls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Xls.Core;

namespace ApplyGradientFillEffects
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
            workbook.Version = ExcelVersion.Version2010;

            //Get the first sheet
            Worksheet sheet = workbook.Worksheets[0];

            //Get "B5" cell
            CellRange range =sheet.Range["B5"];

            //Set row height and column width
            range.RowHeight = 50;
            range.ColumnWidth = 30;
            range.Text = "Hello";

            //Set alignment style
            range.Style.HorizontalAlignment = HorizontalAlignType.Center;

            //Set gradient filling effects
            range.Style.Interior.FillPattern = ExcelPatternType.Gradient;
            range.Style.Interior.Gradient.ForeColor = Color.FromArgb(255, 255, 255);
            range.Style.Interior.Gradient.BackColor = Color.FromArgb(79, 129, 189);
            range.Style.Interior.Gradient.TwoColorGradient(GradientStyleType.Horizontal, GradientVariantsType.ShadingVariants1);

            // Save to
            String result = "ApplyGradientFillEffects_result.xlsx";
            workbook.SaveToFile(result, ExcelVersion.Version2010);

            // Dispose of the workbook object to release resources
            workbook.Dispose();

            // Launch the document
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
