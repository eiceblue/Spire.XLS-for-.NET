Imports System.Collections
Imports System.ComponentModel

Imports Spire.Xls
Imports Spire.Xls.Core

Namespace TillPicAsTextureInShape
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a workbook
			Dim workbook As New Workbook()

			'Load the document from disk
			workbook.LoadFromFile("..\..\..\..\..\..\Data\TillPicAsTextureInShape.xlsx")

			'Get the first worksheet
			Dim sheet As Worksheet = workbook.Worksheets(0)

			'Get the first shape
			Dim shape As IPrstGeomShape = sheet.PrstGeomShapes(0)

			'Fill shape with texture
			shape.Fill.FillType = ShapeFillType.Texture

			'Custom texture with picture
			shape.Fill.CustomTexture("..\..\..\..\..\..\Data\logo.png")

			'Tile pciture as texture 
			shape.Fill.Tile = True

			'Save the document
			Dim output As String = "TillPicAsTextureInShape_out.xlsx"
			workbook.SaveToFile(output, ExcelVersion.Version2013)

			'Launch the Excel file
			ExcelDocViewer(output)
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
