Imports Spire.Xls
Imports Spire.Xls.Core.Spreadsheet.Shapes
Imports System.Drawing.Imaging

Namespace ShapeToImage
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a workbook
			Dim workbook As New Workbook()

			'Load the Excel document from disk
			workbook.LoadFromFile("..\..\..\..\..\..\Data\ShapeToImage.xlsx")

			'Get the first worksheet
			Dim sheet1 As Worksheet = workbook.Worksheets(0)

			'Get the first shape from the first worksheet
			Dim shape As XlsShape = TryCast(sheet1.PrstGeomShapes(0), XlsShape)

			'Save the shape to a image
			Dim img As Image = shape.SaveToImage()
			img.Save("ShapeToImage.png", ImageFormat.Png)
			FileViewer("ShapeToImage.png")
		End Sub

		Private Sub FileViewer(ByVal fileName As String)
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
