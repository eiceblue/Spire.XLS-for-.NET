using System;
using System.Data.OleDb;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.IO;

using Spire.Xls;
using Spire.Xls.Charts;

namespace ReadStream
{

	public partial class Form1 : Form
	{
        public Form1()
        {
            InitializeComponent();
        }
		private void btnRun_Click(object sender, System.EventArgs e)
		{
			Workbook workbook = new Workbook();
			
			//Open excel from a stream
            FileStream fileStream = File.OpenRead(@"..\..\..\..\..\..\Data\ReadStream.xlsx");
			fileStream.Seek(0, SeekOrigin.Begin);

			workbook.LoadFromStream(fileStream);

            workbook.SaveToFile("ReadStream_result.xlsx",ExcelVersion.Version2013);
            ExcelDocViewer("ReadStream_result.xlsx");
		}

		private void ExcelDocViewer( string fileName )
		{
			try
			{
				System.Diagnostics.Process.Start(fileName);
			}
			catch{}
		}

        private void btnClose_Click(object sender, EventArgs e)
        {
            Close();
        }

	}
}
