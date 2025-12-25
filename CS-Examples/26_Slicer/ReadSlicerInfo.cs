using Spire.Xls;
using Spire.Xls.Core;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.IO;

namespace ReadSlicerInfo
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

            StringBuilder builder = new StringBuilder();

            builder.AppendLine("slicers.Count：" + slicers.Count);

            for (int i = 0; i < slicers.Count; i++)
            {
                XlsSlicer xlsSlicer = slicers[i];
                builder.AppendLine();
                builder.AppendLine("xlsSlicer.Name：" + xlsSlicer.Name);
                builder.AppendLine("xlsSlicer.Caption：" + xlsSlicer.Caption);
                builder.AppendLine("xlsSlicer.NumberOfColumns：" + xlsSlicer.NumberOfColumns);
                builder.AppendLine("xlsSlicer.ColumnWidth：" + xlsSlicer.ColumnWidth);
                builder.AppendLine("xlsSlicer.RowHeight：" + xlsSlicer.RowHeight);
                builder.AppendLine("xlsSlicer.ShowCaption：" + xlsSlicer.ShowCaption);
                builder.AppendLine("xlsSlicer.PositionLocked：" + xlsSlicer.PositionLocked);
                builder.AppendLine("xlsSlicer.Width：" + xlsSlicer.Width);
                builder.AppendLine("xlsSlicer.Height：" + xlsSlicer.Height);

                XlsSlicerCache slicerCache = xlsSlicer.SlicerCache;

                builder.AppendLine("slicerCache.SourceName：" + slicerCache.SourceName);
                builder.AppendLine("slicerCache.IsTabular：" + slicerCache.IsTabular);
                builder.AppendLine("slicerCache.Name：" + slicerCache.Name);

                XlsSlicerCacheItemCollection slicerCacheItems = slicerCache.SlicerCacheItems;
                XlsSlicerCacheItem xlsSlicerCacheItem = slicerCacheItems[1];

                builder.AppendLine("xlsSlicerCacheItem.Selected：" + xlsSlicerCacheItem.Selected);
            }

            File.WriteAllText("ReadSlicerInfo.txt", builder.ToString());

            // Dispose of the workbook object to release resources
            wb.Dispose();

            // Launch the file
            ExcelDocViewer("ReadSlicerInfo.txt");
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
