Imports Spire.Xls
Imports System.ComponentModel
Imports System.Text

Namespace CopyOnlyFormulaValue
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			Dim workbook As New Workbook()
			workbook.LoadFromFile("..\..\..\..\..\..\Data\CopyOnlyFormulaValue.xlsx")

			Dim sheet As Worksheet = workbook.Worksheets(0)

			'Set the copy option
			Dim copyOptions As CopyRangeOptions = CopyRangeOptions.OnlyCopyFormulaValue

			'Copy range
			sheet.Copy(sheet.Range("A2:C2"), sheet.Range("A5:C5"), copyOptions)

			'Save to file
			workbook.SaveToFile("result.xlsx", ExcelVersion.Version2010)

			'View the document
			FileViewer("result.xlsx")
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
