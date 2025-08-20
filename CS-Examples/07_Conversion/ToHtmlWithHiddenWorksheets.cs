using Spire.Xls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace ToHtmlWithHiddenWorksheets
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
            Workbook book = new Workbook();

            // Load the document
            book.LoadFromFile(@"..\..\..\..\..\..\Data\ToHtmlWithHiddenWorksheets.xlsx");
			
			// Save Excel to Html
			// false --- To Html with the hidden Worksheet
            // true--- To Html without the hidden Worksheet
            string result = "result.html";
            book.SaveToHtml(result, false);


			// Dispose of the workbook object to release resources
			book.Dispose();
			
            // Launch the document
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
