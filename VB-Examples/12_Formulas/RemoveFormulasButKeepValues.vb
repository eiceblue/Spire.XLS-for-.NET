Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel
Imports Spire.Xls
Imports Spire.Xls.Core.Spreadsheet

Namespace RemoveFormulasButKeepValues
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a workbook.
			Dim workbook As New Workbook()

			'Load the file from disk.
			workbook.LoadFromFile("..\..\..\..\..\..\Data\RemoveFormulasButKeepValues.xlsx")

			'Loop through worksheets.
			For Each sheet As Worksheet In workbook.Worksheets
				'Loop through cells.
				For Each cell As CellRange In sheet.Range
					'If the cell contain formula, get the formula value, clear cell content, and then fill the formula value into the cell.
					If cell.HasFormula Then
						Dim value As Object = cell.FormulaValue
						cell.Clear(ExcelClearOptions.ClearContent)
						cell.Value2 = value
					End If
				Next cell
			Next sheet

			Dim result As String = "Result-RemoveFormulasButKeepValues.xlsx"

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
