Imports System.Collections
Imports System.ComponentModel

Imports Spire.Xls
Imports Spire.Xls.Core

Namespace AddRectangleShape
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a workbook
			Dim workbook As New Workbook()

			'Load the document from disk
			workbook.LoadFromFile("..\..\..\..\..\..\Data\ExcelSample_N1.xlsx")

			'Get the first worksheet
			Dim sheet As Worksheet = workbook.Worksheets(0)

			'Add rectangle shape 1------Rect
			Dim rect1 As IRectangleShape=sheet.RectangleShapes.AddRectangle(11, 2, 60, 100, RectangleShapeType.Rect)
			rect1.Line.Weight = 1
			'Fill shape with solid color
			rect1.Fill.FillType = ShapeFillType.SolidColor
			rect1.Fill.ForeColor = Color.DarkGreen

			'Add rectangle shape 2------RoundRect
			Dim rect2 As IRectangleShape = sheet.RectangleShapes.AddRectangle(11, 5, 60, 100, RectangleShapeType.RoundRect)
			rect2.Line.Weight = 1
			rect2.Fill.FillType = ShapeFillType.SolidColor
			rect2.Fill.ForeColor = Color.DarkCyan

			'Save the document
			Dim output As String = "AddRectangleShape_out.xlsx"
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
