Imports Spire.Xls
Imports Spire.Xls.Core
Imports Spire.Xls.Core.Spreadsheet.PivotTables
Imports System.IO
Imports System.Reflection.Emit

Namespace AddFiltersToFields
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			Dim outputFile As String = "output.xlsx"
			' Create a new workbook object
			Dim workbook As New Workbook()

			'Load the file from disk.
			 workbook.LoadFromFile("..\..\..\..\..\..\Data\PivotTableExample.xlsx")

			'Retrieve the first pivot table from the second sheet
			Dim pt As XlsPivotTable = TryCast(workbook.Worksheets(1).PivotTables(0), XlsPivotTable)

			'Add a label filter to the first row field of the pivot table
			pt.RowFields(0).AddLabelFilter(PivotLabelFilterType.Between, "Argentina", "Nicaragua")

			' Add a value filter on the first row field of the pivot table
			pt.RowFields(0).AddValueFilter(PivotValueFilterType.LessThan, pt.DataFields(0), 5300000, Nothing)

			 '  pt.ColumnFields[0].AddLabelFilter(PivotLabelFilterType.Between, "Argentina", "Nicaragua");
			 '  pt.ColumnFields[0].AddValueFilter(PivotValueFilterType.LessThan, pt.DataFields[0], 5300000, null);

			pt.CalculateData()

			MessageBox.Show(pt.DataFields(0).Name)

			workbook.SaveToFile(outputFile, ExcelVersion.Version2013)

			' Dispose of the workbook object
			workbook.Dispose()

			FileViewer(outputFile)

			Me.Close()
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

		Private Sub label1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles label1.Click

		End Sub

		Private Sub Form1_Load(ByVal sender As Object, ByVal e As EventArgs) Handles MyBase.Load

		End Sub
	End Class
End Namespace
