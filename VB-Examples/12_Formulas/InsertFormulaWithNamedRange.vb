Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel

Imports Spire.Xls
Imports Spire.Xls.Core


Namespace InsertFormulaWithNamedRange
	Partial Public Class Form1
		Inherits Form

		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a workbook
			Dim workbook As New Workbook()
			Dim sheet As Worksheet = workbook.Worksheets(0)

			'Set value
			sheet.Range("A1").Value = "1"
			sheet.Range("A2").Value = "1"

			'Create a named range
			Dim NamedRange As INamedRange = workbook.NameRanges.Add("NewNamedRange")

			NamedRange.NameLocal = "=SUM(A1+A2)"

			'Set the formula
			sheet.Range("C1").Formula = "NewNamedRange"

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
