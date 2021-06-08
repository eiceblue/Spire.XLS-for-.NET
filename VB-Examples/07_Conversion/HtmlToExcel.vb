Imports System.Collections
Imports System.ComponentModel

Imports Spire.Xls
Imports System.Text
Imports System.IO

Namespace HtmlToExcel
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'File path
			Dim filePath As String = "..\..\..\..\..\..\Data\HtmlToExcel.html"

			'Create a workbook
			Dim workbook As New Workbook()

			'Load html
			workbook.LoadFromHtml(filePath)

			'Save to Excel file
			Dim result As String = "HtmlToExcel_result.xlsx"

			workbook.SaveToFile(result, ExcelVersion.Version2013)

			'Launch the file
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
