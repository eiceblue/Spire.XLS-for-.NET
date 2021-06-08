Imports Spire.Xls
Imports System.ComponentModel
Imports System.Text
Imports Spire.Xls.Core.Spreadsheet

Namespace FormatARow
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
			Dim style As CellStyle = workbook.Styles.Add("newStyle")

			'Set the vertical alignment of the text
			style.VerticalAlignment = VerticalAlignType.Center

			'Set the horizontal alignment of the text
			style.HorizontalAlignment = HorizontalAlignType.Center

			'Set the font color of the text
			style.Font.Color = Color.Blue

			'Shrink the text to fit in the cell
			style.ShrinkToFit = True

			'Set the bottom border color of the cell to OrangeRed
			style.Borders(BordersLineType.EdgeBottom).Color = Color.OrangeRed

			'Set the bottom border type of the cell to Dotted
			style.Borders(BordersLineType.EdgeBottom).LineStyle = LineStyleType.Dotted

			'Apply the style to the second row
			sheet.Rows(1).CellStyleName = style.Name

			sheet.Rows(1).Text = "Test"

			Dim result As String = "FormatARow_result.xlsx"

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
