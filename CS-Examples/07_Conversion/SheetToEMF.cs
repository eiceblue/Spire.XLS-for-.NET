using System;
using System.Drawing;
using System.Windows.Forms;

using Spire.Xls;
using System.Drawing.Imaging;
using System.IO;

namespace SheetToEMF
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

            //Load the document from disk
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\ConversionSample1.xlsx");

            //Get the first worksheet in excel workbook
            Worksheet sheet = workbook.Worksheets[0];

            //Create a memory stream
            MemoryStream stream = new MemoryStream();
            
            //Save excel worksheet into EMF stream
            sheet.ToEMFStream(stream, 1, 1, 28, 8, EmfType.EmfPlusDual);

            //Save to file
            Image img = Image.FromStream(stream);
            string output = "ToEMF.emf";
            img.Save(output);

            //Launch the Excel file
            ExcelDocViewer(output);
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
