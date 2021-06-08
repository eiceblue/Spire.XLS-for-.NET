Imports System.Collections
Imports System.ComponentModel

Imports Spire.Xls
Imports Spire.Xls.Core

Namespace SetFormulaWithNamedRange
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a workbook 
			Dim workbook As New Workbook()

			'Create an empty sheet
			workbook.LoadFromFile("..\..\..\..\..\..\Data\ExcelSample_N1.xlsx")

			'Get the sheet
			Dim sheet As Worksheet = workbook.Worksheets(0)

			'Create a named range
			Dim NamedRange As INamedRange = workbook.NameRanges.Add("MyNamedRange")
			'Refers to range
			NamedRange.RefersToRange = sheet.Range("B10:B12")

			'Set the formula of range to named range
			sheet.Range("B13").Formula = "=SUM(MyNamedRange)"

			'Set value of ranges
			sheet.Range("B10").Value2=10
			sheet.Range("B11").Value2 = 20
			sheet.Range("B12").Value2 = 30

			'Save the Excel file
			Dim result As String = "SetFormulaWithNamedRange_out.xlsx"
			workbook.SaveToFile(result, ExcelVersion.Version2013)

			'Launch the Excel file
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
