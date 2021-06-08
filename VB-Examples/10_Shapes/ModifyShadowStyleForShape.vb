Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel
Imports Spire.Xls
Imports Spire.Xls.Core.Spreadsheet
Imports Spire.Xls.Core

Namespace ModifyShadowStyleForShape
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a workbook.
			Dim workbook As New Workbook()

			'Load the file from disk.
			workbook.LoadFromFile("..\..\..\..\..\..\Data\Template_Xls_5.xlsx")

			'Get the first worksheet.
			Dim sheet As Worksheet = workbook.Worksheets(0)

			'Get the third shape from the worksheet.
			Dim shape As IPrstGeomShape = sheet.PrstGeomShapes(2)

			'Set the shadow style for the shape.
			shape.Shadow.Angle = 90
			shape.Shadow.Transparency = 30
			shape.Shadow.Distance = 10
			shape.Shadow.Size = 130
			shape.Shadow.Color = Color.Yellow
			shape.Shadow.Blur = 30
			shape.Shadow.HasCustomStyle = True

			Dim result As String = "Result-ModifyShadowStyleForShape.xlsx"

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
