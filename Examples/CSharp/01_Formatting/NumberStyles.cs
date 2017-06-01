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
			this.label1.Size = new System.Drawing.Size(495, 18);
			this.label1.TabIndex = 4;
			this.label1.Text = "The sample demonstrates how to set number formatting in an excel workbook.";
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

			sheet.Range["B1"].Text = "NUMBER FORMATTING";
			sheet.Range["B1"].Style.Font.IsBold = true;

			sheet.Range["B3"].Text = "0";
			sheet.Range["C3"].NumberValue = 1234.5678;
			sheet.Range["C3"].NumberFormat = "0";

			sheet.Range["B4"].Text = "0.00";
			sheet.Range["C4"].NumberValue = 1234.5678;
			sheet.Range["C4"].NumberFormat = "0.00";

			sheet.Range["B5"].Text = "#,##0.00";
			sheet.Range["C5"].NumberValue = 1234.5678;
			sheet.Range["C5"].NumberFormat = "#,##0.00";

			sheet.Range["B6"].Text = "$#,##0.00";
			sheet.Range["C6"].NumberValue = 1234.5678;
			sheet.Range["C6"].NumberFormat = "$#,##0.00";

			sheet.Range["B7"].Text = "0;[Red]-0";
			sheet.Range["C7"].NumberValue = -1234.5678;
			sheet.Range["C7"].NumberFormat = "0;[Red]-0";

			sheet.Range["B8"].Text = "0.00;[Red]-0.00";
			sheet.Range["C8"].NumberValue = -1234.5678;
			sheet.Range["C8"].NumberFormat = "0.00;[Red]-0.00";

			sheet.Range["B9"].Text = "#,##0;[Red]-#,##0";
			sheet.Range["C9"].NumberValue = -1234.5678;
			sheet.Range["C9"].NumberFormat = "#,##0;[Red]-#,##0";

			sheet.Range["B10"].Text = "#,##0.00;[Red]-#,##0.000";
			sheet.Range["C10"].NumberValue = -1234.5678;
			sheet.Range["C10"].NumberFormat = "#,##0.00;[Red]-#,##0.00";

			sheet.Range["B11"].Text = "0.00E+00";
			sheet.Range["C11"].NumberValue = 1234.5678;
			sheet.Range["C11"].NumberFormat = "0.00E+00";

			sheet.Range["B12"].Text = "0.00%";
			sheet.Range["C12"].NumberValue = 1234.5678;
			sheet.Range["C12"].NumberFormat = "0.00%";

			sheet.Range["B3:B12"].Style.KnownColor = ExcelColors.Gray25Percent;

			
			sheet.AutoFitColumn(2);
			sheet.AutoFitColumn(3);
			
		
			workbook.SaveToFile("Sample.xls");
			ExcelDocViewer(workbook.FileName);
		}

		private void btnAbout_Click(object sender, System.EventArgs e)
		{
			Close();
		}


	}
}
