using System;
using System.Windows.Forms;
using Spire.Xls;

namespace SparkLine
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

            //Load a Workbook from disk
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\SparkLine.xlsx");

            //Get the first sheet
            Worksheet sheet = workbook.Worksheets[0];

            //Add sparkline
            SparklineGroup sparklineGroup = sheet.SparklineGroups.AddGroup(SparklineType.Line);
            SparklineCollection sparklines = sparklineGroup.Add();
            sparklines.Add(sheet["A2:D2"], sheet["E2"]);
            sparklines.Add(sheet["A3:D3"], sheet["E3"]);
            sparklines.Add(sheet["A4:D4"], sheet["E4"]);
            sparklines.Add(sheet["A5:D5"], sheet["E5"]);
            sparklines.Add(sheet["A6:D6"], sheet["E6"]);
            sparklines.Add(sheet["A7:D7"], sheet["E7"]);
            sparklines.Add(sheet["A8:D8"], sheet["E8"]);
            sparklines.Add(sheet["A9:D9"], sheet["E9"]);
            sparklines.Add(sheet["A10:D10"], sheet["E10"]);
            sparklines.Add(sheet["A11:D11"], sheet["E11"]);

            //Save the file
            workbook.SaveToFile("Output.xlsx",ExcelVersion.Version2010);

            // Dispose of the workbook object to release resources
            workbook.Dispose();

            // Launch the file
            ExcelDocViewer("Output.xlsx");
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
