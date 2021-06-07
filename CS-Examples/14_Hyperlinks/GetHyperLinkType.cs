using System;
using System.IO;
using System.Text;
using System.Windows.Forms;
using Spire.Xls;

namespace GetHyperLinkType
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
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\HyperlinksSample2.xlsx");

            //Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            //Iterate all hyperlinks
            StringBuilder sb = new StringBuilder();
            foreach (var item in sheet.HyperLinks)
            {
                //Get hyperlink address
                string address = item.Address;
                //Get hyperlink type
                HyperLinkType type = item.Type;
                sb.AppendLine("Link address: " + address);
                sb.AppendLine("Link type: " + type.ToString());
                sb.AppendLine();
            }

            //Save to Text file
            string output = "GetHyperLinkType.txt";
            File.WriteAllText(output, sb.ToString());

            //Launch the file
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
