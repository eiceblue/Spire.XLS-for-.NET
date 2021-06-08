Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel

Imports Spire.Xls

Namespace LinkToOtherSheetCell
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

			Dim range As CellRange = sheet.Range("A1")

			'Add hyperlink in the range
			Dim hyperlink As HyperLink = sheet.HyperLinks.Add(range)

			'Set the link type
			hyperlink.Type = HyperLinkType.Workbook

			'Set the display text
			hyperlink.TextToDisplay = "Link to Sheet2 cell C5"

			'Set the address
			hyperlink.Address = "Sheet2!C5"

			'Save and launch result file
			Dim result As String = "result.xlsx"
			workbook.SaveToFile(result, ExcelVersion.Version2010)
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
