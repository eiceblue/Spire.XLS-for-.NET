Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel
Imports Spire.Xls
Imports Spire.Xls.Core.Spreadsheet
Imports Spire.Xls.Core

Namespace SetShadowStyleForShape
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

            ' Add an ellipse shape to the worksheet at position (5, 5) with size 150x100
            Dim ellipse As IPrstGeomShape = sheet.PrstGeomShapes.AddPrstGeomShape(5, 5, 150, 100, PrstGeomShapeType.Ellipse)

            ' Set the shadow style of the ellipse shape
            ellipse.Shadow.Angle = 90
            ellipse.Shadow.Distance = 10
            ellipse.Shadow.Size = 150
            ellipse.Shadow.Color = Color.Gray
            ellipse.Shadow.Blur = 30
            ellipse.Shadow.Transparency = 1
            ellipse.Shadow.HasCustomStyle = True

            ' Specify the file name for the resulting Excel file
            Dim result As String = "Result-SetShapeShadowStyle.xlsx"

            ' Save the workbook to the specified file in Excel 2013 format
            workbook.SaveToFile(result, ExcelVersion.Version2013)
            ' Release the resources used by the workbook
            workbook.Dispose()

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
