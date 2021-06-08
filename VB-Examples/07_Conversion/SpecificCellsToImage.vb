Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel

Imports Spire.Xls
Imports System.Drawing.Imaging

Namespace SpecificCellsToImage

	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a workbook
			Dim workbook As New Workbook()

			'Load the document from disk
			workbook.LoadFromFile("..\..\..\..\..\..\Data\ConversionSample1.xlsx")

			'Get the first worksheet in Excel file
			Dim sheet As Worksheet = workbook.Worksheets(0)

			'Specify Cell Ranges and Save to certain Image formats
			sheet.ToImage(1, 1, 7, 5).Save("image1.png", ImageFormat.Png)
			sheet.ToImage(8, 1, 15, 5).Save("image2.jpg", ImageFormat.Jpeg)
			sheet.ToImage(17, 1, 23, 5).Save("image3.bmp", ImageFormat.Bmp)
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
