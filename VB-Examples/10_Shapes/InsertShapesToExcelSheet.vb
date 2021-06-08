Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel
Imports Spire.Xls
Imports Spire.Xls.Core.Spreadsheet
Imports Spire.Xls.Core

Namespace InsertShapesToExcelSheet
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a workbook.
			Dim workbook As New Workbook()

			'Get the first worksheet.
			Dim sheet As Worksheet = workbook.Worksheets(0)

			'Add a triangle shape.
			Dim triangle As IPrstGeomShape = sheet.PrstGeomShapes.AddPrstGeomShape(2, 2, 100, 100, PrstGeomShapeType.Triangle)
			'Fill the triangle with solid color.
			triangle.Fill.ForeColor = Color.Yellow
			triangle.Fill.FillType = ShapeFillType.SolidColor

			'Add a heart shape.
			Dim heart As IPrstGeomShape = sheet.PrstGeomShapes.AddPrstGeomShape(2, 5, 100, 100, PrstGeomShapeType.Heart)
			'Fill the heart with gradient color.
			heart.Fill.ForeColor = Color.Red
			heart.Fill.FillType = ShapeFillType.Gradient

			'Add an arrow shape with default color.
			Dim arrow As IPrstGeomShape = sheet.PrstGeomShapes.AddPrstGeomShape(10, 2, 100, 100, PrstGeomShapeType.CurvedRightArrow)

			'Add a cloud shape.
			Dim cloud As IPrstGeomShape = sheet.PrstGeomShapes.AddPrstGeomShape(10, 5, 100, 100, PrstGeomShapeType.Cloud)
			'Fill the cloud with custom picture
			cloud.Fill.CustomPicture(Image.FromFile("..\..\..\..\..\..\Data\SpireXls.png"), "SpireXls.png")
			cloud.Fill.FillType = ShapeFillType.Picture

			Dim result As String = "Result-InsertShapesToExcelSheet.xlsx"

			'Save to file.
			workbook.SaveToFile(result, ExcelVersion.Version2013)

			'Launch the MS Excel file.
			ExcelDocViewer(result)
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
