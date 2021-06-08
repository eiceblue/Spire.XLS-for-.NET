Imports Spire.Xls
Imports System.ComponentModel
Imports System.Text
Imports Spire.Xls.Core

Namespace ApplyGradientFillEffects
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a workbook
			Dim workbook As New Workbook()
			workbook.Version = ExcelVersion.Version2010
			'Get the first sheet
			Dim sheet As Worksheet = workbook.Worksheets(0)
			'Get "B5" cell
			Dim range As CellRange =sheet.Range("B5")
			'Set row height and column width
			range.RowHeight = 50
			range.ColumnWidth = 30
			range.Text = "Hello"

			'Set alignment style
			range.Style.HorizontalAlignment = HorizontalAlignType.Center

			'Set gradient filling effects
			range.Style.Interior.FillPattern = ExcelPatternType.Gradient
			range.Style.Interior.Gradient.ForeColor = Color.FromArgb(255, 255, 255)
			range.Style.Interior.Gradient.BackColor = Color.FromArgb(79, 129, 189)
			range.Style.Interior.Gradient.TwoColorGradient(GradientStyleType.Horizontal, GradientVariantsType.ShadingVariants1)

			Dim result As String = "ApplyGradientFillEffects_result.xlsx"

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
