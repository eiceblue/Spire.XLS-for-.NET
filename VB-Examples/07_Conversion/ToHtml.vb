Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel

Imports Spire.Xls
Imports Spire.Xls.Core.Spreadsheet

Namespace ToHtml
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			Dim workbook As New Workbook()

			workbook.LoadFromFile("..\..\..\..\..\..\Data\ToHtml.xlsx")

			Dim sheet As Worksheet = workbook.Worksheets(0)
			Dim options As New HTMLOptions()
			options.ImageEmbedded = True
			sheet.SaveToHtml("sample.html")
			ExcelDocViewer("sample.html")
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
