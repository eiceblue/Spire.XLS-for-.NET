using Spire.Xls;
using Spire.Xls.Core;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace ModifySlicer
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            // Create a new Workbook instance
            Workbook wb = new Workbook();

            // Load an existing Excel file from the specified path
            wb.LoadFromFile(@"..\..\..\..\..\..\Data\SlicerTemplate.xlsx");

            // Get the first worksheet in the workbook
            Worksheet worksheet = wb.Worksheets[0];

            // Get the slicer collection from the worksheet
            XlsSlicerCollection slicers = worksheet.Slicers;

            // Get the first slicer from the slicer collection
            XlsSlicer xlsSlicer = slicers[0];

            // Set the style of the slicer to a dark theme (style type 4)
            xlsSlicer.StyleType = SlicerStyleType.SlicerStyleDark4;

            // Change the caption (title) of the slicer
            xlsSlicer.Caption = "Modified Slicer";

            // Lock the position of the slicer to prevent it from being moved in the worksheet
            xlsSlicer.PositionLocked = true;

            // Get the collection of cache items associated with the slicer
            XlsSlicerCacheItemCollection slicerCacheItems = xlsSlicer.SlicerCache.SlicerCacheItems;

            // Get the first cache item in the collection
            XlsSlicerCacheItem xlsSlicerCacheItem = slicerCacheItems[0];

            // Deselect the cache item
            xlsSlicerCacheItem.Selected = false;

            // Get the display value of the cache item
            string displayValue = xlsSlicerCacheItem.DisplayValue;

            // Get the slicer cache associated with the slicer
            XlsSlicerCache slicerCache = xlsSlicer.SlicerCache;

            // Set the cross-filter type to show items even if they have no associated data
            slicerCache.CrossFilterType = SlicerCacheCrossFilterType.ShowItemsWithNoData;

            // Save the modified workbook to a new file with Excel 2013 version format
            wb.SaveToFile("ModifySlicer.xlsx", ExcelVersion.Version2013);

            // Dispose of the workbook object to release resources
            wb.Dispose();

            // Launch the file
            ExcelDocViewer("ModifySlicer.xlsx");
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
