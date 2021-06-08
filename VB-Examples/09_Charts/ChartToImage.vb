Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel
Imports System.Drawing.Imaging
Imports Spire.Xls

Namespace ChartToImage
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Load a Workbook from disk
			Dim workbook As New Workbook()
			workbook.LoadFromFile("..\..\..\..\..\..\Data\ChartToImage.xlsx")

			'Save chart as image
		   Dim image As Image= workbook.SaveChartAsImage(workbook.Worksheets(0), 0)

		   image.Save("Output.png",ImageFormat.Png)

			'Launch the file
			ExcelDocViewer("Output.png")
		End Sub
		Private Sub ExcelDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub

		Private Sub btnClose_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnClose.Click
			Close()
		End Sub
	End Class
End Namespace
