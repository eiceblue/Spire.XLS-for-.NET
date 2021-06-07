using Spire.Xls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Xls.Core.Spreadsheet.PivotTables;

namespace RepeatItemLabels
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
            //Load an excel file including pivot table
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\RepeatItemLabelsExample.xlsx");
            //Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            //Add an empty worksheet 
            Worksheet sheet2 = workbook.CreateEmptySheet();
            //Add PivotTable
            sheet2.Name = "Pivot Table";
            CellRange dataRange = sheet.Range["A1:D9"];
            PivotCache cache = workbook.PivotCaches.Add(dataRange);
            PivotTable pt = sheet2.PivotTables.Add("Pivot Table", sheet.Range["A1"], cache);
            var r1 = pt.PivotFields["VendorNo"];
            r1.Axis = AxisTypes.Row;
            pt.Options.RowHeaderCaption = "VendorNo";
            r1.Subtotals = SubtotalTypes.None;

            r1.RepeatItemLabels = true;
            //Repeat item lables
            pt.PivotFields["OnHand"].RepeatItemLabels = true;
            pt.Options.RowLayout = PivotTableLayoutType.Tabular;
            var r2 = pt.PivotFields["Desc"];
            r2.Axis = AxisTypes.Row;
            pt.DataFields.Add(pt.PivotFields["OnHand"], "Sum of onHand", SubtotalTypes.None);
            pt.BuiltInStyle = PivotBuiltInStyles.PivotStyleMedium12;
            String result = "RepeatItemLabels_result.xlsx";
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

        private void Form1_Load(object sender, EventArgs e)
        {

        }
    }
}
