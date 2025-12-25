using Spire.Xls;
using Spire.Xls.Core;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace CreateSlicerFromPivotTable
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            Workbook wb = new Workbook();
            Worksheet worksheet = wb.Worksheets[0];
            worksheet.Range["A1"].Value = "fruit";
            worksheet.Range["A2"].Value = "grape";
            worksheet.Range["A3"].Value = "blueberry";
            worksheet.Range["A4"].Value = "kiwi";
            worksheet.Range["A5"].Value = "cherry";
            worksheet.Range["A6"].Value = "grape";
            worksheet.Range["A7"].Value = "blueberry";
            worksheet.Range["A8"].Value = "kiwi";
            worksheet.Range["A9"].Value = "cherry";

            worksheet.Range["B1"].Value = "year";
            worksheet.Range["B2"].Value2 = 2020;
            worksheet.Range["B3"].Value2 = 2020;
            worksheet.Range["B4"].Value2 = 2020;
            worksheet.Range["B5"].Value2 = 2020;
            worksheet.Range["B6"].Value2 = 2021;
            worksheet.Range["B7"].Value2 = 2021;
            worksheet.Range["B8"].Value2 = 2021;
            worksheet.Range["B9"].Value2 = 2021;

            worksheet.Range["C1"].Value = "amount";
            worksheet.Range["C2"].Value2 = 50;
            worksheet.Range["C3"].Value2 = 60;
            worksheet.Range["C4"].Value2 = 70;
            worksheet.Range["C5"].Value2 = 80;
            worksheet.Range["C6"].Value2 = 90;
            worksheet.Range["C7"].Value2 = 100;
            worksheet.Range["C8"].Value2 = 110;
            worksheet.Range["C9"].Value2 = 120;

            // Get pivot table collection
            Spire.Xls.Collections.PivotTablesCollection pivotTables = worksheet.PivotTables;

            //Add a PivotTable to the worksheet
            CellRange dataRange = worksheet.Range["A1:C9"];
            PivotCache cache = wb.PivotCaches.Add(dataRange);

            //Cell to put the pivot table
            Spire.Xls.PivotTable pt = worksheet.PivotTables.Add("TestPivotTable", worksheet.Range["A12"], cache);

            //Drag the fields to the row area.
            PivotField pf = pt.PivotFields["fruit"] as PivotField;
            pf.Axis = AxisTypes.Row;
            PivotField pf2 = pt.PivotFields["year"] as PivotField;
            pf2.Axis = AxisTypes.Column;

            //Drag the field to the data area.
            pt.DataFields.Add(pt.PivotFields["amount"], "SUM of Count", SubtotalTypes.Sum);

            //Set PivotTable style
            pt.BuiltInStyle = PivotBuiltInStyles.PivotStyleMedium10;

            pt.CalculateData();

            //Get slicer collection
            XlsSlicerCollection slicers = worksheet.Slicers;

            int index = slicers.Add(pt, "E12", 0);

            XlsSlicer xlsSlicer = slicers[index];
            xlsSlicer.Name = "xlsSlicer";
            xlsSlicer.Width = 100;
            xlsSlicer.Height = 120;
            xlsSlicer.StyleType = SlicerStyleType.SlicerStyleLight2;
            xlsSlicer.PositionLocked = true;

            //Get SlicerCache object of current slicer
            XlsSlicerCache slicerCache = xlsSlicer.SlicerCache;
            slicerCache.CrossFilterType = SlicerCacheCrossFilterType.ShowItemsWithNoData;

            //Style setting
            XlsSlicerCacheItemCollection slicerCacheItems = xlsSlicer.SlicerCache.SlicerCacheItems;
            XlsSlicerCacheItem xlsSlicerCacheItem = slicerCacheItems[0];
            xlsSlicerCacheItem.Selected = false;

            XlsSlicerCollection slicers_2 = worksheet.Slicers;

            IPivotField r1 = pt.PivotFields["year"];
            int index_2 = slicers_2.Add(pt, "I12", r1);

            XlsSlicer xlsSlicer_2 = slicers[index_2];
            xlsSlicer_2.RowHeight = 40;
            xlsSlicer_2.StyleType = SlicerStyleType.SlicerStyleLight3;
            xlsSlicer_2.PositionLocked = false;

            //Get SlicerCache object of current slicer
            XlsSlicerCache slicerCache_2 = xlsSlicer_2.SlicerCache;
            slicerCache_2.CrossFilterType = SlicerCacheCrossFilterType.ShowItemsWithDataAtTop;

            //Style setting
            XlsSlicerCacheItemCollection slicerCacheItems_2 = xlsSlicer_2.SlicerCache.SlicerCacheItems;
            XlsSlicerCacheItem xlsSlicerCacheItem_2 = slicerCacheItems_2[1];
            xlsSlicerCacheItem_2.Selected = false;
            pt.CalculateData();

            //Save to file
            wb.SaveToFile("CreateSlicerFromPivotTable.xlsx", ExcelVersion.Version2013);

            // Dispose of the workbook object to release resources
            wb.Dispose();

            // Launch the file
            ExcelDocViewer("CreateSlicerFromPivotTable.xlsx");
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
