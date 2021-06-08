Imports Spire.Xls
Imports System.ComponentModel
Imports System.Text

Namespace UseArrayFormulas
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a workbook
			Dim workbook As New Workbook()

			'Get the first sheet
			Dim sheet As Worksheet = workbook.Worksheets(0)

			sheet.Range("A1").NumberValue = 1
			sheet.Range("A2").NumberValue = 2
			sheet.Range("A3").NumberValue = 3
			sheet.Range("B1").NumberValue = 4
			sheet.Range("B2").NumberValue = 5
			sheet.Range("B3").NumberValue = 6
			sheet.Range("C1").NumberValue = 7
			sheet.Range("C2").NumberValue = 8
			sheet.Range("C3").NumberValue = 9

			'Write array formula
			sheet.Range("A5:C6").FormulaArray="=LINEST(A1:A3,B1:C3,TRUE,TRUE)"

			'Calculate Formulas
			workbook.CalculateAllValue()

			Dim result As String = "UseArrayFormulas_result.xlsx"

			'Save to file
			workbook.SaveToFile(result, ExcelVersion.Version2010)

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
