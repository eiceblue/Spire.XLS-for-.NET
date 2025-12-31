Imports System
Imports System.Drawing
Imports System.Windows.Forms
Imports System.Drawing.Imaging
Imports Spire.Xls

Namespace ChartToImage
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As System.EventArgs)
			' Create a workbook
			Dim workbook As New Workbook()

			'Load file from disk
			workbook.LoadFromFile("..\..\..\..\..\..\Data\ChartToImage.xlsx")

			'Save chart as image
			Dim image As Image= workbook.SaveChartAsImage(workbook.Worksheets(0), 0)
			image.Save("Output.png",ImageFormat.Png)

			'////////////////Use the following code for netstandard dlls/////////////////////////
'			
'            Stream image = workbook.SaveChartAsImage(workbook.Worksheets[0], 0);
'            string filename = String.Format("ChartToImage.png");
'            FileStream fileStream = new FileStream(filename, FileMode.Create, FileAccess.Write);
'            image.CopyTo(fileStream, 100);
'            fileStream.Flush();
'            fileStream.Close();
'            image.Close();
'			

			' Dispose of the workbook object to release resources
			workbook.Dispose()

			'Launch the file
			ExcelDocViewer("Output.png")
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
