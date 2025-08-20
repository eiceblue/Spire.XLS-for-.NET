using System;
using System.Data.OleDb;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;

using Spire.Xls;
using Spire.Xls.Core;

namespace InsertOLEObjects
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        private void btnRun_Click(object sender, System.EventArgs e)
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet in the workbook
            Worksheet ws = workbook.Worksheets[0];

            // Set the text in cell A1
            ws.Range["A1"].Text = "Here is an OLE Object.";

            // Insert an OLE object
            string xlsFile = @"..\..\..\..\..\..\Data\InsertOLEObjects.xls";
            Image image = GenerateImage(xlsFile);
            IOleObject oleObject = ws.OleObjects.Add(xlsFile, image, OleLinkType.Embed);
            oleObject.Location = ws.Range["B4"];
            oleObject.ObjectType = OleObjectType.ExcelWorksheet;

            // Specify the output file name for the result
            string result = "InsertOLEObjects_result.xlsx";

            // Save the modified workbook to the specified file using Excel 2010 format
            workbook.SaveToFile(result, ExcelVersion.Version2010);

            // Dispose of the workbook object to release resources
            workbook.Dispose();

            // Launch the file
            ExcelDocViewer(result);
        }
        private Image GenerateImage(string fileName)
        {
            Workbook book = new Workbook();
            book.LoadFromFile(fileName);
            book.Worksheets[0].PageSetup.LeftMargin = 0;
            book.Worksheets[0].PageSetup.RightMargin = 0;
            book.Worksheets[0].PageSetup.TopMargin = 0;
            book.Worksheets[0].PageSetup.BottomMargin = 0;
            return book.Worksheets[0].ToImage(1, 1, 19, 5);
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
