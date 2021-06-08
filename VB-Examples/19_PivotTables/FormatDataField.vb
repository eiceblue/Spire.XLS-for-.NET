Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel

Imports Spire.Xls
Imports Spire.Xls.Core.Spreadsheet.PivotTables

Namespace FormatDataField
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			Dim workbook As New Workbook()
			workbook.LoadFromFile("..\..\..\..\..\..\Data\FormatDataField.xlsx")
			Dim sheet As Worksheet = workbook.Worksheets(0)
			' Access the PivotTable
			Dim pt As XlsPivotTable = TryCast(sheet.PivotTables(0), XlsPivotTable)
			' Access the data field.
			Dim pivotDataField As PivotDataField = pt.DataFields(0)
			' Set data display format
			pivotDataField.ShowDataAs = PivotFieldFormatType.PercentageOfColumn

			Dim result As String = "FormatDataField_output.xlsx"
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
