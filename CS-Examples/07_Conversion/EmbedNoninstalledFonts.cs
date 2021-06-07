using System;
using System.Data.OleDb;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using Spire.Xls;
using Spire.Xls.Charts;

namespace EmbedNoninstalledFonts
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
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\EmbedNoninstalledFonts.xlsx");

            //Get the first sheet
            Worksheet sheet = workbook.Worksheets[0];

            //Get the first chart
            Chart chart = sheet.Charts[0];

            //Load the font file from disk
            workbook.CustomFontFilePaths = new string[] { @"..\..\..\..\..\..\Data\PT_Serif-Caption-Web-Regular.ttf" };
            System.Collections.Hashtable result = workbook.GetCustomFontParsedResult(); 
    
            ArrayList valueList = new ArrayList(result.Values);
  
            //Apply the font for PrimaryValueAxis of chart
            chart.PrimaryValueAxis.Font.FontName = valueList[0] as string;

            //Apply the font for PrimaryCategoryAxis of chart
            chart.PrimaryCategoryAxis.Font.FontName = valueList[0] as string;

            //Apply the font for the first chartSerie of chart
            ChartSerie chartSerie1 = chart.Series[0];
            chartSerie1.DataPoints.DefaultDataPoint.DataLabels.FontName = valueList[0] as string;

            string output ="Output.pdf";
            //Save and Launch
            workbook.SaveToFile(output, Spire.Xls.FileFormat.PDF);
            ExcelDocViewer(output);

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
