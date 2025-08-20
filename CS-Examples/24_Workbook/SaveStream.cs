using System;
using System.Data.OleDb;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.IO;

using Spire.Xls;

namespace SaveStream
{
	public partial class Form1 : Form
	{
        public Form1()
        {
            InitializeComponent();
        }
		private void btnRun_Click(object sender, System.EventArgs e)
		{
            // Create a workbook
            Workbook workbook = new Workbook();

            // Load the document from disk
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\SaveStream.xls");

			// Save an excel workbook to stream
            FileStream fileStream = new FileStream("SaveStream.xlsx", FileMode.Create);
			workbook.SaveToStream(fileStream,FileFormat.Version2010);

            // Close the stream
            fileStream.Close();

            // Dispose of the workbook object to release resources 
            workbook.Dispose();

            // Launch the file
            ExcelDocViewer("SaveStream.xlsx");
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
