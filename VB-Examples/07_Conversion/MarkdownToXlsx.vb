Imports Spire.Xls
Imports System.ComponentModel
Imports System.Text

Namespace MarkdownToXlsx
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			' Create a new Workbook instance
			Dim workbook As New Workbook()

			' Load content from a Markdown file into the workbook
			workbook.LoadFromMarkdown("..\..\..\..\..\..\Data\sample.md")

			' Define the output file name for the saved Excel file
			Dim result As String = "MarkdownToXlsx.xlsx"

			' Save the workbook to a file in Excel 2016 format (.xlsx)
			workbook.SaveToFile(result, ExcelVersion.Version2016)

			' Release the resources used by the workbook object
			workbook.Dispose()
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
