using System;
using System.Data.OleDb;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;

using Spire.Xls;
using System.IO;

namespace ToSVG
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
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\ToSVG.xlsx");
            for (int i = 0; i < workbook.Worksheets.Count; i++)
            {
                FileStream fs = new FileStream(string.Format("sheet{0}.svg", i), FileMode.Create);
                workbook.Worksheets[i].ToSVGStream(fs, 0, 0, 0, 0);
                fs.Flush();
                fs.Close();
            }
			 System.Diagnostics.Process.Start("sheet0.svg");
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            Close();
        }
    }
}
