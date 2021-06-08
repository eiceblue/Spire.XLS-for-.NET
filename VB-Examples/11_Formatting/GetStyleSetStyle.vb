Imports Spire.Xls
Imports System.ComponentModel
Imports System.Text
Imports Spire.Xls.Core.Spreadsheet

Namespace GetStyleSetStyle
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a workbook
			Dim workbook As New Workbook()

			'Load a excel document
			workbook.LoadFromFile("..\..\..\..\..\..\Data\templateAz.xlsx")

			'Get the first sheet
			Dim sheet As Worksheet = workbook.Worksheets(0)

			'Get "B4" cell
			Dim range As CellRange = sheet.Range("B4")
			'Get the style of cell
			Dim style As CellStyle = range.Style
			style.Font.FontName = "Calibri"
			style.Font.IsBold = True
			style.Font.Size = 15
			style.Font.Color = Color.CornflowerBlue

			range.Style = style

			Dim result As String = "UseGetStyleSetStyle_result.xlsx"

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
