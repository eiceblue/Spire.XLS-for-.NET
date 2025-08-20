using System;
using System.Collections.Generic;
using System.Windows.Forms;
using Spire.Xls;

namespace AddCustomObject
{

	public partial class Form1 : Form
	{
        public class Student
        {
            internal Student(string name, int age)
            {
                this.Name = name;
                this.Age = age;
            }
            public string Name { get; set; }
            public int Age { get; set; }
        }

        public Form1()
        {
            InitializeComponent();
        }

		private void btnRun_Click(object sender, EventArgs e)
		{
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet in the workbook
            Worksheet sheet = workbook.Worksheets[0];

            // Set marker designer field in cell A1
            sheet.Range["A1"].Value = "&=Student.Name";
            sheet.Range["B1"].Value = "&=Student.Age";

            // Create a list of Student objects
            List<Student> list = new List<Student>();
            list.Add(new Student("John", 16));
            list.Add(new Student("Mary", 17));
            list.Add(new Student("Lucy", 17));

            // Fill custom object using the "Student" parameter
            workbook.MarkerDesigner.AddParameter("Student", list);
            workbook.MarkerDesigner.Apply();
            workbook.CalculateAllValue();

            // AutoFit rows and columns to adjust their sizes based on content
            sheet.AllocatedRange.AutoFitRows();
            sheet.AllocatedRange.AutoFitColumns();

            // Save the document to a file named "AddCustomObject.xlsx" using Excel 2013 format
            string output = "AddCustomObject.xlsx";
            workbook.SaveToFile(output, ExcelVersion.Version2013);

            // Dispose of the workbook object to release resources
            workbook.Dispose();

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
