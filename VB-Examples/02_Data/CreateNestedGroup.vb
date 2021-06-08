Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel
Imports Spire.Xls
Imports Spire.Xls.Core.Spreadsheet

Namespace CreateNestedGroup
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a workbook.
			Dim workbook As New Workbook()

			'Get the first worksheet.
			Dim sheet As Worksheet = workbook.Worksheets(0)

			'Set the style.
			Dim style As CellStyle = workbook.Styles.Add("style")
			style.Font.Color = Color.CadetBlue
			style.Font.IsBold = True

			'Set the summary rows appear above detail rows.
			sheet.PageSetup.IsSummaryRowBelow = False

			'Insert sample data to cells.
			sheet.Range("A1").Value = "Project plan for project X"
			sheet.Range("A1").CellStyleName = style.Name

			sheet.Range("A3").Value = "Set up"
			sheet.Range("A3").CellStyleName = style.Name
			sheet.Range("A4").Value = "Task 1"
			sheet.Range("A5").Value = "Task 2"
			sheet.Range("A4:A5").BorderAround(LineStyleType.Thin)
			sheet.Range("A4:A5").BorderInside(LineStyleType.Thin)

			sheet.Range("A7").Value = "Launch"
			sheet.Range("A7").CellStyleName = style.Name
			sheet.Range("A8").Value = "Task 1"
			sheet.Range("A9").Value = "Task 2"
			sheet.Range("A8:A9").BorderAround(LineStyleType.Thin)
			sheet.Range("A8:A9").BorderInside(LineStyleType.Thin)

			'Group the rows that you want to group.
			sheet.GroupByRows(2, 9, False)
			sheet.GroupByRows(4, 5, False)
			sheet.GroupByRows(8, 9, False)

			Dim result As String = "Result-CreateNestedGroup.xlsx"

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
