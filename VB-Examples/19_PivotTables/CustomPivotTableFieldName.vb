Imports Spire.Xls
Imports System.ComponentModel
Imports System.Text
Imports Spire.Xls.Core.Spreadsheet.PivotTables

Namespace CustomPivotTableFieldName
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			' Create a workbook
			Dim workbook As New Workbook()

			' Load an excel file including pivot table   
			workbook.LoadFromFile("..\..\..\..\..\..\Data\CustomPivotTableFieldName.xlsx")

			' Get the sheet in which the pivot table is located
			Dim sheet As Worksheet = workbook.Worksheets("PivotTable")

			' Access the first pivot table in the worksheet
			Dim pivotTable As XlsPivotTable = TryCast(sheet.PivotTables(0), XlsPivotTable)

			' Set a custom name for the row field
			pivotTable.RowFields(0).CustomName = "custom_rowName"

			' Set a custom name for the column field
			pivotTable.ColumnFields(0).CustomName = "custom_colName"

			' Set a custom name for the data field
			pivotTable.DataFields(0).CustomName = "custom_DataName"

			' Calculate the pivot table data
			pivotTable.CalculateData()

			' Specify the filename for the resulting workbook
			Dim result As String = "CustomPivotTableFieldName_result.xlsx"

			' Save the modified workbook to a file
			workbook.SaveToFile(result, ExcelVersion.Version2010)

			' Dispose of the workbook object to release resources
			workbook.Dispose()

			'View the document
			FileViewer(result)
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
