using Spire.Xls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace ReplaceFont
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            // Create a workbook
            Workbook workbook = new Workbook();

            // Load a excel document
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\CreateTable.xlsx");

            // Get the second sheet
            Worksheet sheet = workbook.Worksheets[0];

            // Define the new style
            CellStyle newStyle = workbook.Styles.Add("newStyle");
            newStyle.Font.FontName = "Arial Black";
            newStyle.Font.Size = 14;

            // The old style which need to be replaced
            CellStyle oldStyle = null;

            for (int i = 0; i < workbook.Styles.Count; i++)
            {
                if (workbook.Styles[i].Font.FontName == "Aleo")
                {
                   oldStyle = sheet.Range["D9"].Style;            
                }
            }

            // Replace style
            sheet.ReplaceAll("North America", oldStyle, "America", newStyle);

            // Save the file
            workbook.SaveToFile("ReplaceFont_out.xlsx", ExcelVersion.Version2013);

            // Dispose of the workbook object to free up resources
            workbook.Dispose();

            // Launch the document
            ExcelDocViewer("ReplaceFont_out.xlsx");
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
