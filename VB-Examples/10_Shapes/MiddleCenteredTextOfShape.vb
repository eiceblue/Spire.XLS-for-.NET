Imports Spire.Xls
Imports Spire.Xls.Core
Imports System.ComponentModel
Imports System.Text

Namespace MiddleCenteredTextOfShape
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new Workbook object
            Dim workbook As New Workbook()

            ' Get the first worksheet from the workbook
            Dim sheet As Worksheet = workbook.Worksheets(0)

            ' Add a rectangle shape to the worksheet at position (8, 2) with size 300x300
            Dim rect As IPrstGeomShape = sheet.PrstGeomShapes.AddPrstGeomShape(8, 2, 300, 300, PrstGeomShapeType.Rect)

            ' Set the fill color of the rectangle shape to white
            rect.Fill.ForeColor = Color.White
            ' Set the fill type of the rectangle shape to solid color
            rect.Fill.FillType = ShapeFillType.SolidColor

            ' Set the text content of the rectangle shape to "E-iceblue"
            rect.Text = "E-iceblue"
            ' Set the vertical alignment of the text within the rectangle shape to centered
            rect.TextVerticalAlignment = ExcelVerticalAlignment.MiddleCentered

            ' Save the workbook to a file named "result.xlsx" in Excel 2013 format
            workbook.SaveToFile("result.xlsx", ExcelVersion.Version2013)
            ' Release the resources used by the workbook
            workbook.Dispose()

            'View the document
            FileViewer("result.xlsx")
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
