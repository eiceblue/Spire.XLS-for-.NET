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
            ' Declare and initialize a new Workbook object
            Dim workbook As New Workbook()

            ' Load an Excel file from the specified path
            workbook.LoadFromFile("..\..\..\..\..\..\Data\ExcelSample_N1.xlsx")

            ' Get the first worksheet from the workbook
            Dim sheet As Worksheet = workbook.Worksheets(0)

            ' Add a rectangle shape to the worksheet at position (11, 2) with width 60, height 100, and rectangle shape type
            Dim rect1 As IRectangleShape = sheet.RectangleShapes.AddRectangle(11, 2, 60, 100, RectangleShapeType.Rect)
            ' Set the weight of the line to 1
            rect1.Line.Weight = 1

            ' Set the fill type of rect1 to solid color and set the foreground color to DarkGreen
            rect1.Fill.FillType = ShapeFillType.SolidColor
            rect1.Fill.ForeColor = Color.DarkGreen

            ' Add another rectangle shape to the worksheet at position (11, 5) with width 60, height 100, and round rectangle shape type
            Dim rect2 As IRectangleShape = sheet.RectangleShapes.AddRectangle(11, 5, 60, 100, RectangleShapeType.RoundRect)
            ' Set the weight of the line to 1
            rect2.Line.Weight = 1
            ' Set the fill type of rect2 to solid color
            rect2.Fill.FillType = ShapeFillType.SolidColor
            ' Set the foreground color of rect2 to DarkCyan
            rect2.Fill.ForeColor = Color.DarkCyan

            ' Specify the output file name
            Dim output As String = "AddRectangleShape_out.xlsx"

            ' Save the modified workbook to the specified file in Excel 2013 format
            workbook.SaveToFile(output, ExcelVersion.Version2013)
            ' Release the resources used by the workbook
            workbook.Dispose()

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
