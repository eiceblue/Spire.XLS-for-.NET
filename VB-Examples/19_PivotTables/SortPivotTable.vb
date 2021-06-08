Imports Spire.Xls
Imports System.ComponentModel
Imports System.Text

Namespace SortPivotTable
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a workbook
			Dim workbook As New Workbook()
			'Load an excel file including pivot table
			workbook.LoadFromFile("..\..\..\..\..\..\Data\SortPivotTable.xlsx")
			'Get the first worksheet
			Dim sheet As Worksheet = workbook.Worksheets(0)
			'Add an empty worksheet 
			Dim sheet2 As Worksheet = workbook.CreateEmptySheet()

			sheet2.Name = "Pivot Table"
			'Specify the datasorce
			Dim dataRange As CellRange = sheet.Range("A1:C9")
			Dim cache As PivotCache = workbook.PivotCaches.Add(dataRange)
			'Add PivotTable
			Dim pt As PivotTable = sheet2.PivotTables.Add("Pivot Table", sheet.Range("A1"), cache)
			Dim r1 As PivotField = TryCast(pt.PivotFields("No"), PivotField)
			r1.Axis = AxisTypes.Row
			pt.Options.RowLayout = PivotTableLayoutType.Tabular
			'Sort PivotField
			r1.SortType = PivotFieldSortType.Descending

			Dim r2 As PivotField = TryCast(pt.PivotFields("Name"), PivotField)
			r2.Axis = AxisTypes.Row
			pt.DataFields.Add(pt.PivotFields("OnHand"), "Sum of onHand", SubtotalTypes.None)
			pt.BuiltInStyle = PivotBuiltInStyles.PivotStyleMedium12

			Dim result As String = "SortPivotTable_result.xlsx"
			'Save to file
			workbook.SaveToFile(result, ExcelVersion.Version2013)

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
