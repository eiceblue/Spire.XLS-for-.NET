using Spire.Xls;
using System;
using System.Windows.Forms;


namespace CustomPaperSizeForPrinting
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            // Create a workbook
            Workbook workbook = new Workbook();

            // Load file from disk
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\CustomPaperSizeForPrinting.xlsx");

            // Get the first worksheet 
            Worksheet worksheet = workbook.Worksheets[0];

            // Set the paper size to the printer's custom paper size
            worksheet.PageSetup.CustomPaperSizeName = "customPaper";

            //Custom the paper size directly
            //sheet.PageSetup.SetCustomPaperSize(224, (float)50);

            //Set the page orientation
            //sheet.PageSetup.Orientation = PageOrientationType.Portrait;

            // Use the default printer to print
            workbook.PrintDocument.Print();

            // Dispose of the workbook object to release resources
            workbook.Dispose();
        }     

        private void btnClose_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }
    }
}
