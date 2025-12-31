Imports System
Imports System.Windows.Forms
Imports System.Drawing
Imports Spire.Xls

Namespace FillChartElementWithPicture

	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs)
			'Create a workbook
			Dim workbook As New Workbook()

			'Load the document from disk
			workbook.LoadFromFile("..\..\..\..\..\..\Data\ChartSample1.xlsx")

			'Get the first worksheet from workbook
			Dim ws As Worksheet = workbook.Worksheets(0)
			'Get the first chart
			Dim chart As Chart = ws.Charts(0)

			' A. Fill chart area with image
			chart.ChartArea.Fill.CustomPicture(Image.FromFile("..\..\..\..\..\..\Data\background.png"), "None")
			chart.PlotArea.Fill.Transparency = 0.9

			'////////////////Use the following code for netstandard dlls/////////////////////////
'            
'            FileStream fs = new FileStream(@"..\..\..\..\..\..\Data\background.png", FileMode.Open, FileAccess.Read, FileShare.Read);
'            byte[] bytes = new byte[fs.Length];
'            fs.Read(bytes, 0, bytes.Length);
'            fs.Close();
'            Stream ImgFile1 = new MemoryStream(bytes);
'            chart.ChartArea.Fill.CustomPicture(ImgFile1, "None");
'			

			'// B.Fill plot area with image
			'chart.PlotArea.Fill.CustomPicture(Image.FromFile(@"..\..\..\..\..\..\Data\background.png"), "None");

			'////////////////Use the following code for netstandard dlls/////////////////////////
'            
'            FileStream fs = new FileStream(@"..\..\..\..\..\..\Data\background.png", FileMode.Open, FileAccess.Read, FileShare.Read);
'            byte[] bytes = new byte[fs.Length];
'            fs.Read(bytes, 0, bytes.Length);
'            fs.Close();
'            Stream ImgFile2 = new MemoryStream(bytes);
'            chart.PlotArea.Fill.CustomPicture(ImgFile2, "None");
'			

			'Save the document
			Dim output As String = "FillChartElementWithPicture.xlsx"
			workbook.SaveToFile(output, ExcelVersion.Version2010)

			' Dispose of the workbook object to release resources
			workbook.Dispose()

			'Launch the file
			ExcelDocViewer(output)
		End Sub
		Private Sub ExcelDocViewer(ByVal fileName As String)
			Try
				System.Diagnostics.Process.Start(fileName)
			Catch
			End Try
		End Sub

		Private Sub btnClose_Click(ByVal sender As Object, ByVal e As EventArgs)
			Close()
		End Sub

	End Class
End Namespace
