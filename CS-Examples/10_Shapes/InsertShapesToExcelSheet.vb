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
            ' Create a new Workbook object
            Dim workbook As New Workbook()

            ' Get the first worksheet from the workbook
            Dim sheet As Worksheet = workbook.Worksheets(0)

            ' Add a triangle shape to the worksheet at position (2, 2) with size 100x100
            Dim triangle As IPrstGeomShape = sheet.PrstGeomShapes.AddPrstGeomShape(2, 2, 100, 100, PrstGeomShapeType.Triangle)
            triangle.Fill.ForeColor = Color.Yellow
            triangle.Fill.FillType = ShapeFillType.SolidColor

            ' Add a heart shape to the worksheet at position (2, 5) with size 100x100
            Dim heart As IPrstGeomShape = sheet.PrstGeomShapes.AddPrstGeomShape(2, 5, 100, 100, PrstGeomShapeType.Heart)
            heart.Fill.ForeColor = Color.Red
            heart.Fill.FillType = ShapeFillType.Gradient

            ' Add a curved right arrow shape to the worksheet at position (10, 2) with size 100x100
            Dim arrow As IPrstGeomShape = sheet.PrstGeomShapes.AddPrstGeomShape(10, 2, 100, 100, PrstGeomShapeType.CurvedRightArrow)

            ' Add a cloud shape to the worksheet at position (10, 5) with size 100x100
            Dim cloud As IPrstGeomShape = sheet.PrstGeomShapes.AddPrstGeomShape(10, 5, 100, 100, PrstGeomShapeType.Cloud)
            cloud.Fill.CustomPicture(Image.FromFile("..\..\..\..\..\..\Data\SpireXls.png"), "SpireXls.png")
            cloud.Fill.FillType = ShapeFillType.Picture

            ' Specify the file name for the resulting Excel file
            Dim result As String = "Result-InsertShapesToExcelSheet.xlsx"

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
