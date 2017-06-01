using System;
using System.Data.OleDb;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;

using Spire.Xls;

namespace Spire.Xls.Sample
{
	/// <summary>
	/// Summary description for Form1.
	/// </summary>
	public class Form1 : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Button btnRun;
		private System.Windows.Forms.Button btnAbout;
		private System.Windows.Forms.Label label1;
		/// <summary>
		/// Required designer variable.
		/// </summary
		private System.ComponentModel.Container components = null;

		public Form1()
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();
			//
			// TODO: Add any constructor code after InitializeComponent call
			//
		}

		/// <summary>
		/// Clean up any resources being used.
		/// </summary>
		protected override void Dispose( bool disposing )
		{
			if( disposing )
			{
				if (components != null) 
				{
					components.Dispose();
				}
			}
			base.Dispose( disposing );
		}

		#region Windows Form Designer generated code
		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
			this.btnRun = new System.Windows.Forms.Button();
			this.btnAbout = new System.Windows.Forms.Button();
			this.label1 = new System.Windows.Forms.Label();
			this.SuspendLayout();
			// 
			// btnRun
			// 
			this.btnRun.Location = new System.Drawing.Point(448, 56);
			this.btnRun.Name = "btnRun";
			this.btnRun.Size = new System.Drawing.Size(72, 23);
			this.btnRun.TabIndex = 2;
			this.btnRun.Text = "Run";
			this.btnRun.Click += new System.EventHandler(this.btnRun_Click);
			// 
			// btnAbout
			// 
			this.btnAbout.Location = new System.Drawing.Point(528, 56);
			this.btnAbout.Name = "btnAbout";
			this.btnAbout.TabIndex = 3;
			this.btnAbout.Text = "Close";
			this.btnAbout.Click += new System.EventHandler(this.btnAbout_Click);
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Font = new System.Drawing.Font("Tahoma", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(134)));
			this.label1.Location = new System.Drawing.Point(16, 16);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(468, 18);
			this.label1.TabIndex = 4;
			this.label1.Text = "The sample demonstrates how to create conditional formatting in an excel workbook.";
			// 
			// Form1
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(6, 14);
			this.ClientSize = new System.Drawing.Size(616, 94);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.btnAbout);
			this.Controls.Add(this.btnRun);
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "Form1";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Spire.XLS sample";
			this.ResumeLayout(false);

		}
		#endregion

		/// <summary>
		/// The main entry point for the application.
		/// </summary>
		[STAThread]
		static void Main() 
		{
			Application.Run(new Form1());
		}

		private void ExcelDocViewer( string fileName )
		{
			try
			{
				System.Diagnostics.Process.Start(fileName);
			}
			catch{}
		}

		private void btnRun_Click(object sender, System.EventArgs e)
		{
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Range["A1"].NumberValue = 582;
            sheet.Range["A2"].NumberValue = 234;
            sheet.Range["A3"].NumberValue = 314;
            sheet.Range["A4"].NumberValue = 50;
            sheet.Range["A5"].NumberValue = 100;
            sheet.Range["B1"].NumberValue = 150;
            sheet.Range["B2"].NumberValue = 894;
            sheet.Range["B3"].NumberValue = 560;
            sheet.Range["B4"].NumberValue = 900;
            sheet.Range["B5"].NumberValue = 70;
            sheet.Range["C1"].NumberValue = 134;
            sheet.Range["C2"].NumberValue = 700;
            sheet.Range["C3"].NumberValue = 920;
            sheet.Range["C4"].NumberValue = 450;
            sheet.Range["C5"].NumberValue = 50;
            sheet.AllocatedRange.RowHeight = 15;
            sheet.AllocatedRange.ColumnWidth = 17;

            //create conditional formatting rule
            ConditionalFormatWrapper format1 = sheet.Range["A1:C1"].ConditionalFormats.AddCondition();
            format1.FormatType = ConditionalFormatType.CellValue;
            format1.FirstFormula = "150";
            format1.Operator = ComparisonOperatorType.Greater;
            format1.FontColor = Color.Red;
            format1.BackColor = Color.LightSalmon;

            ConditionalFormatWrapper format2 = sheet.Range["A2:C2"].ConditionalFormats.AddCondition();
            format2.FormatType = ConditionalFormatType.CellValue;
            format2.FirstFormula = "300";
            format2.Operator = ComparisonOperatorType.Less;
            format2.FontColor = Color.Green;
            format2.BackColor = Color.LightBlue;

            //add data bars
            ConditionalFormatWrapper format3 = sheet.Range["A3:C3"].ConditionalFormats.AddCondition();
            format3.FormatType = ConditionalFormatType.DataBar;
            format3.DataBar.BarColor = Color.CadetBlue;

            //add icon sets
            ConditionalFormatWrapper format4 = sheet.Range["A4:C4"].ConditionalFormats.AddCondition();
            format4.FormatType = ConditionalFormatType.IconSet;
            format4.IconSet.IconSetType = IconSetType.FiveArrows;

            //add color scales
            ConditionalFormatWrapper format5 = sheet.Range["A5:C5"].ConditionalFormats.AddCondition();
            format5.FormatType = ConditionalFormatType.ColorScale;


            workbook.SaveToFile("Sample.xlsx", ExcelVersion.Version2010);
			ExcelDocViewer(workbook.FileName);
		}

		private void btnAbout_Click(object sender, System.EventArgs e)
		{
			Close();
		}


	}
}
