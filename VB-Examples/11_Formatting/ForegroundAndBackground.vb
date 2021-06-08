Imports Spire.Xls
Imports System.ComponentModel
Imports System.Text
Imports Spire.Xls.Core.Spreadsheet

Namespace ForegroundAndBackground
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

			'Create a new style
			Dim style As CellStyle = workbook.Styles.Add("newStyle1")

			'Set filling pattern type
			style.Interior.FillPattern = ExcelPatternType.VerticalStripe

			'Set filling Background color
			style.Interior.Gradient.BackKnownColor = ExcelColors.Green

			'Set filling Foreground color
			style.Interior.Gradient.ForeKnownColor = ExcelColors.Yellow

			'Apply the style to  "B2" cell
			sheet.Range("B2").CellStyleName = style.Name
			sheet.Range("B2").Text = "Test"
			sheet.Range("B2").RowHeight = 30
			sheet.Range("B2").ColumnWidth = 50


			'Create a new style
			style = workbook.Styles.Add("newStyle2")

			'Set filling pattern type
			style.Interior.FillPattern = ExcelPatternType.ThinHorizontalStripe

			'Set filling Foreground color
			style.Interior.Gradient.ForeKnownColor = ExcelColors.Red

			'Apply the style to  "B4" cell
			sheet.Range("B4").CellStyleName = style.Name
			sheet.Range("B4").RowHeight = 30
			sheet.Range("B4").ColumnWidth = 60

			Dim result As String = "ForegroundAndBackground_result.xlsx"

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
