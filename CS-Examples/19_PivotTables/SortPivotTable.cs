using Spire.Xls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace SortPivotTable
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
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\SortPivotTable.xlsx");
            //Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            //Add an empty worksheet 
            Worksheet sheet2 = workbook.CreateEmptySheet();

            sheet2.Name = "Pivot Table";
            //Specify the datasorce
            CellRange dataRange = sheet.Range["A1:C9"];
            PivotCache cache = workbook.PivotCaches.Add(dataRange);
            //Add PivotTable
            PivotTable pt = sheet2.PivotTables.Add("Pivot Table", sheet.Range["A1"], cache);
            PivotField r1 = pt.PivotFields["No"] as PivotField;
            r1.Axis = AxisTypes.Row;
            pt.Options.RowLayout = PivotTableLayoutType.Tabular;
            //Sort PivotField
            r1.SortType = PivotFieldSortType.Descending;

            PivotField r2 = pt.PivotFields["Name"] as PivotField;
            r2.Axis = AxisTypes.Row;
            pt.DataFields.Add(pt.PivotFields["OnHand"], "Sum of onHand", SubtotalTypes.None);
            pt.BuiltInStyle = PivotBuiltInStyles.PivotStyleMedium12;

            String result = "SortPivotTable_result.xlsx";
            //Save to file
            workbook.SaveToFile(result, ExcelVersion.Version2013);

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
