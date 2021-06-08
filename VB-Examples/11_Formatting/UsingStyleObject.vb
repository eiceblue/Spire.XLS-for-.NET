Imports Spire.Xls
Imports System.ComponentModel
Imports System.Text
Imports Spire.Xls.Core.Spreadsheet

Namespace UsingStyleObject
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a workbook
			Dim workbook As New Workbook()

			'Add a new worksheet to the Excel object
			Dim sheet As Worksheet = workbook.Worksheets.Add("new sheet")

			'Access the "B1" cell from the worksheet
			Dim cell As CellRange = sheet.Range("B1")

			'Add some value to the "B1" cell
			cell.Text = "Hello Spire!"

			'Create a new style
			Dim style As CellStyle = workbook.Styles.Add("newStyle")

			'Set the vertical alignment of the text in the "B1" cell
			style.VerticalAlignment = VerticalAlignType.Center

			'Set the horizontal alignment of the text in the "B1" cell
			style.HorizontalAlignment = HorizontalAlignType.Center

			'Set the font color of the text in the "B1" cell
			style.Font.Color = Color.Blue

			'Shrink the text to fit in the cell
			style.ShrinkToFit = True

			'Set the bottom border color of the cell to GreenYellow
			style.Borders(BordersLineType.EdgeBottom).Color = Color.GreenYellow

			'Set the bottom border type of the cell to Medium
			style.Borders(BordersLineType.EdgeBottom).LineStyle = LineStyleType.Medium

			'Assign the Style object to the "B1" cell
			cell.Style = style


			'Apply the same style to some other cells
			sheet.Range("B4").Style = style
			sheet.Range("B4").Text = "Test"
			sheet.Range("C3").CellStyleName = style.Name
			sheet.Range("C3").Text = "Welcome to use Spire.XLS"
			sheet.Range("D4").Style = style

			Dim result As String = "UsingStyleObject_result.xlsx"

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
