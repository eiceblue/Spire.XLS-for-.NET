using Spire.Xls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.IO;

namespace GetListOfFontsUsed
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

            //Load a excel document
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\templateAz.xlsx");
            
            List<ExcelFont> fonts = new List<ExcelFont>();

            //Loop all sheets of workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                for (int r = 0; r < sheet.Rows.Length; r++)
                {
                    for (int c = 0; c < sheet.Rows[r].CellList.Count; c++)
                    {
                        //Get the font of cell and add it to list
                        fonts.Add(sheet.Rows[r].CellList[c].Style.Font);
                    }
                }
            }
            StringBuilder strB = new StringBuilder();

            foreach (ExcelFont font in fonts)
            {
                strB.AppendLine(String.Format("FontName:{0}; FontSize:{1}",font.FontName,font.Size));
            }

            String result = "GetListOfFontsUsed_result.txt";

            File.WriteAllText(result, strB.ToString());
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
