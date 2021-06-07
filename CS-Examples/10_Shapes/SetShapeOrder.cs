using Spire.Xls;
using Spire.Xls.Core.Spreadsheet.Shapes;
using System;
using System.Windows.Forms;


namespace SetShapeOrder
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
            Workbook wb = new Workbook();
            //Load an excel file
            wb.LoadFromFile(@"..\..\..\..\..\..\Data\SetShapeOrder.xlsx");

            //Bring the picture forward one level
            wb.Worksheets[0].Pictures[0].ChangeLayer(ShapeLayerChangeType.BringForward);

            //Bring the image in fron of all other objects
            wb.Worksheets[1].Pictures[0].ChangeLayer(ShapeLayerChangeType.BringToFront);

            //Send the shape back one level
            XlsShape shape = wb.Worksheets[2].PrstGeomShapes[1] as XlsShape;
            shape.ChangeLayer(ShapeLayerChangeType.SendBackward);

            //Send the shape behind all other objects
            shape = wb.Worksheets[3].PrstGeomShapes[1] as XlsShape;
            shape.ChangeLayer(ShapeLayerChangeType.SendToBack);

            String result = "SetShapeOrder_result.xlsx";
            //Save to file
            wb.SaveToFile(result, ExcelVersion.Version2010);
            //View the document
            FileViewer(result);
        }

        private void FileViewer(string fileName)
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
