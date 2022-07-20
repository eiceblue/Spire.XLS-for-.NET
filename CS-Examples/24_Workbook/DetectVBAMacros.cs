using System;
using System.Windows.Forms;
using Spire.Xls;

namespace DetectVBAMacros
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
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\MacroSample.xls");

            //Detect if the Excel file contains VBA macros
            bool hasMacros = false;
            hasMacros = workbook.HasMacros;
            if (hasMacros)
            {
                this.textBox1.Text = "Yes";
            }

            else
            {
                this.textBox1.Text = "No";
            }
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            Close();
        }

	}
}
