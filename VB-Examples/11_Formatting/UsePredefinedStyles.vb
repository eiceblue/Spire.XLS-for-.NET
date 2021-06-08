Imports Spire.Xls
Imports System.ComponentModel
Imports System.Text
Imports Spire.Xls.Core.Spreadsheet

Namespace UsePredefinedStyles
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
			style.Font.FontName = "Calibri"
			style.Font.IsBold = True
			style.Font.Size = 15
			style.Font.Color = Color.CornflowerBlue

			'Get "B5" cell
			Dim range As CellRange =sheet.Range("B5")
			range.Text = "Welcome to use Spire.XLS"
			range.CellStyleName = style.Name
			range.AutoFitColumns()

			Dim result As String = "UsePredefinedStyles_result.xlsx"

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
