using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;

using Spire.Xls;
using System.IO;
using System.Text;

namespace OpenFiles
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        private void btnRun_Click(object sender, System.EventArgs e)
        {
             // Pathes of files
            string filepath = @"..\..\..\..\..\..\Data\ExcelSample_N1.xlsx";
            string filepath97 = @"..\..\..\..\..\..\Data\ExcelSample97_N.xls";
            string filepathXml = @"..\..\..\..\..\..\Data\OfficeOpenXML_N.xml";
            string filepathCsv = @"..\..\..\..\..\..\Data\CSVSample_N.csv";

            // Create string builder
            StringBuilder builder = new StringBuilder();

            // 1. Load file by file path
            // Create a workbook
            Workbook workbook1 = new Workbook();
            // Load the document from disk
            workbook1.LoadFromFile(filepath);
            builder.AppendLine("Workbook opened using file path successfully!");

            // Dispose of the workbook object to release resources 
            workbook1.Dispose();

            // 2. Load file by file stream
            FileStream stream = new FileStream(filepath, FileMode.Open);

            // Create a workbook
            Workbook workbook2 = new Workbook();

            // Load the document from disk
            workbook2.LoadFromStream(stream);
            builder.AppendLine("Workbook opened using file stream successfully!");
            stream.Dispose();

            // Dispose of the workbook object to release resources 
            workbook2.Dispose();

            // 3. Open Microsoft Excel 97 - 2003 file
            Workbook wbExcel97 = new Workbook();
            wbExcel97.LoadFromFile(filepath97, ExcelVersion.Version97to2003);
            builder.AppendLine("Microsoft Excel 97 - 2003 workbook opened successfully!");

            // 4. Open xml file
            Workbook wbXML = new Workbook();
            wbXML.LoadFromXml(filepathXml);
            builder.AppendLine("XML file opened successfully!");

            // Dispose of the workbook object to release resources 
            wbExcel97.Dispose();

            //5. Open csv file
            Workbook wbCSV = new Workbook();
            wbCSV.LoadFromFile(filepathCsv, ",", 1, 1);
            builder.AppendLine("CSV file opened successfully!");

            //Save to txt file
            string output = "OpenFiles_out.txt";
            File.WriteAllText(output, builder.ToString());

            // Dispose of the workbook object to release resources 
            wbCSV.Dispose();

            // Launch the file
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
