Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel

Imports Spire.Xls

Namespace SetDefaultRowAndColumnStyle
	Partial Public Class Form1
		Inherits Form

		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click

			Dim workbook As New Workbook()

			'Get the first sheet
			Dim sheet As Worksheet = workbook.Worksheets(0)

			'Create a cell style and set the color
			Dim style As CellStyle = workbook.Styles.Add("Mystyle")
			style.Color = Color.Yellow

			'Set the default style for the first row and column 
			sheet.SetDefaultRowStyle(1, style)
			sheet.SetDefaultColumnStyle(1, style)

			'Save and launch result file
			Dim result As String = "result.xlsx"
			workbook.SaveToFile(result, ExcelVersion.Version2010)
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
