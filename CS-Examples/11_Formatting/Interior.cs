using System;
using System.Data.OleDb;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;

using Spire.Xls;

namespace Interior
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

            //Initialize the workbook
            Worksheet sheet = workbook.Worksheets[0];

            //Specify the version
            workbook.Version = ExcelVersion.Version2007;

            //Define the number of the colors
            int maxColor = Enum.GetValues(typeof(ExcelColors)).Length;

            //Create a random object
            Random random = new Random(10000000);

            for (int i = 2; i < 40; i++)
            {
                //Random backKnownColor
                ExcelColors backKnownColor = (ExcelColors)(random.Next(1, maxColor / 2));

                //Add text
                sheet.Range["A1"].Text = "Color Name";
                sheet.Range["B1"].Text = "Red";
                sheet.Range["C1"].Text = "Green";
                sheet.Range["D1"].Text = "Blue";

                //Merge the sheet"E1-K1"
                sheet.Range["E1:K1"].Merge();
                sheet.Range["E1:K1"].Text = "Gradient";
                sheet.Range["A1:K1"].Style.Font.IsBold = true;
                sheet.Range["A1:K1"].Style.Font.Size = 11;

                //Set the text of color in sheetA-sheetD
                string colorName = backKnownColor.ToString();
                sheet.Range[string.Format("A{0}", i)].Text = colorName;
                sheet.Range[string.Format("B{0}", i)].NumberValue = workbook.GetPaletteColor(backKnownColor).R;
                sheet.Range[string.Format("C{0}", i)].NumberValue = workbook.GetPaletteColor(backKnownColor).G;
                sheet.Range[string.Format("D{0}", i)].NumberValue = workbook.GetPaletteColor(backKnownColor).B;

                //Merge the sheets
                sheet.Range[string.Format("E{0}:K{0}", i)].Merge();

                //Set the text of sheetE-sheetK
                sheet.Range[string.Format("E{0}:K{0}", i)].Text = colorName;

                //Set the interior of the color
                sheet.Range[string.Format("E{0}:K{0}", i)].Style.Interior.FillPattern = ExcelPatternType.Gradient;
                sheet.Range[string.Format("E{0}:K{0}", i)].Style.Interior.Gradient.BackKnownColor = backKnownColor;
                sheet.Range[string.Format("E{0}:K{0}", i)].Style.Interior.Gradient.ForeKnownColor = ExcelColors.White;
                sheet.Range[string.Format("E{0}:K{0}", i)].Style.Interior.Gradient.GradientStyle = GradientStyleType.Vertical;
                sheet.Range[string.Format("E{0}:K{0}", i)].Style.Interior.Gradient.GradientVariant = GradientVariantsType.ShadingVariants1;
            }

            //AutoFit Column
            sheet.AutoFitColumn(1);

            String result = "Result-Interior.xlsx";
            //Save and Launch
            workbook.SaveToFile(result, ExcelVersion.Version2013);
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
