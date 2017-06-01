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
            this.btnAbout.Size = new System.Drawing.Size(75, 23);
            this.btnAbout.TabIndex = 3;
            this.btnAbout.Text = "Close";
            this.btnAbout.Click += new System.EventHandler(this.btnAbout_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Tahoma", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label1.Location = new System.Drawing.Point(16, 16);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(494, 14);
            this.label1.TabIndex = 4;
            this.label1.Text = "The sample demonstrates how to set interior formatting in an excel workbook.";
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
            this.PerformLayout();

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

            workbook.Version = ExcelVersion.Version2007;


            int maxColor = Enum.GetValues(typeof(ExcelColors)).Length;
            Random random = new Random((int)System.DateTime.Now.Ticks);
            for (int i = 2; i < 40; i++)
            {
                
                ExcelColors backKnownColor = (ExcelColors)(random.Next(1,maxColor / 2));

                sheet.Range["A1"].Text = "Color Name";
                sheet.Range["B1"].Text = "Red";
                sheet.Range["C1"].Text = "Green";
                sheet.Range["D1"].Text = "Blue";
                sheet.Range["E1:K1"].Merge();
                sheet.Range["E1:K1"].Text = "Gradient";
                sheet.Range["A1:K1"].Style.Font.IsBold = true;
                sheet.Range["A1:K1"].Style.Font.Size = 11;

                string colorName = backKnownColor.ToString();
                sheet.Range[string.Format("A{0}",i)].Text = colorName;
                sheet.Range[string.Format("B{0}", i)].Text = workbook.GetPaletteColor(backKnownColor).R.ToString();
                sheet.Range[string.Format("C{0}", i)].Text = workbook.GetPaletteColor(backKnownColor).G.ToString();
                sheet.Range[string.Format("D{0}", i)].Text = workbook.GetPaletteColor(backKnownColor).B.ToString();

                sheet.Range[string.Format("E{0}:K{0}",i)].Merge();
                sheet.Range[string.Format("E{0}:K{0}", i)].Text = colorName;
                sheet.Range[string.Format("E{0}:K{0}", i)].Style.Interior.FillPattern = ExcelPatternType.Gradient;
                sheet.Range[string.Format("E{0}:K{0}", i)].Style.Interior.Gradient.BackKnownColor = backKnownColor;
                sheet.Range[string.Format("E{0}:K{0}", i)].Style.Interior.Gradient.ForeKnownColor = ExcelColors.White;
                sheet.Range[string.Format("E{0}:K{0}", i)].Style.Interior.Gradient.GradientStyle = GradientStyleType.Vertical;
                sheet.Range[string.Format("E{0}:K{0}", i)].Style.Interior.Gradient.GradientVariant = GradientVariantsType.ShadingVariants1;
            }

            sheet.AutoFitColumn(1);
		
			workbook.SaveToFile("Sample.xlsx");

			ExcelDocViewer(workbook.FileName);
		}

		private void btnAbout_Click(object sender, System.EventArgs e)
		{
			Close();
		}


	}
}
