using System;
using System.Drawing;
using System.Windows.Forms;
using Spire.Xls;
using Spire.Xls.Core.Spreadsheet.PivotTables;

namespace SetPivotFieldsConditionalFormat
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        private void btnRun_Click(object sender, System.EventArgs e)
        {
            // Create a new workbook object
            Workbook workbook = new Workbook();

            // Load the Excel document from disk
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\PivotTableExample.xlsx");

            // Get the worksheet with the PivotTable
            Worksheet worksheet = workbook.Worksheets["PivotTable"];

            // Get the PivotTable from the worksheet
            PivotTable table = (PivotTable)worksheet.PivotTables[0];

            // Add a conditional format to the PivotTable
            PivotConditionalFormatCollection pcfs = table.PivotConditionalFormats;
            PivotConditionalFormat pc = pcfs.AddPivotConditionalFormat(table.DataFields[0]);
            Spire.Xls.Core.IConditionalFormat cf = pc.AddCondition();
            cf.FormatType = ConditionalFormatType.NotContainsBlanks;
            cf.FillPattern = ExcelPatternType.Solid;
            cf.BackColor = Color.Yellow;

            // Save the modified workbook to a file
            workbook.SaveToFile("output.xlsx", ExcelVersion.Version2016);

            // Dispose of the workbook object to release resources
            workbook.Dispose();

            //Launch the txt file
            ExcelDocViewer("output.xlsx");
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
