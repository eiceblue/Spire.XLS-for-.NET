Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel

Imports Spire.Xls
Imports Spire.Xls.Core.Spreadsheet.PivotTables
Imports Spire.Xls.Core

Namespace GroupPivotTableByDate
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			' Create a new Workbook object
			Dim workbook As New Workbook()

			' Load the workbook from the specified file path
			workbook.LoadFromFile("..\..\..\..\..\..\Data\GroupPivotTableByDate.xlsx")

			' Get the first worksheet in the workbook
			Dim sheet As Worksheet = workbook.Worksheets(0)

			' Get the first pivot table in the worksheet
			Dim pt As XlsPivotTable = TryCast(sheet.PivotTables(0), XlsPivotTable)

			' Get the first row field in the pivot table
			Dim field As IPivotField = pt.RowFields(0)

			' Set the start and end dates for grouping
			Dim start As New Date(2023, 1, 5)
			Dim [end] As New Date(2023, 3, 2)

			' Set the group by type to days
			Dim types() As PivotGroupByTypes = { PivotGroupByTypes.Days }

			' Create a new group with the specified start and end dates, group by type, and interval
			field.CreateGroup(start, [end], types, 10)

			' Calculate the pivot table data
			pt.CalculateData()

			' Refresh the pivot table cache
			pt.Cache.IsRefreshOnLoad = True

			' Set the output file name
			Dim result As String = "GroupPivotTableByDate_output.xlsx"

			' Save the workbook to the specified file path with the specified Excel version
			workbook.SaveToFile(result, ExcelVersion.Version2010)

			' Dispose the workbook object to release resources
			workbook.Dispose()

			' Launch the file
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
