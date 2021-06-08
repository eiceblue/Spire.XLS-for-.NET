Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel
Imports Spire.Xls
Imports Spire.Xls.Core.Spreadsheet
Imports System.Text
Imports System.IO

Namespace RetrieveExternalFileHyperlinks
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a workbook.
			Dim workbook As New Workbook()

			'Load the file from disk.
			workbook.LoadFromFile("..\..\..\..\..\..\Data\RetrieveExternalFileHyperlinks.xlsx")

			'Get the first worksheet.
			Dim sheet As Worksheet = workbook.Worksheets(0)

			Dim content As New StringBuilder()

			'Retrieve external file hyperlinks.
			For Each item As HyperLink In sheet.HyperLinks
				Dim address As String = item.Address
				Dim sheetName As String = item.Range.WorksheetName
				Dim range As CellRange = item.Range
				content.AppendLine(String.Format("Cell[{0},{1}] in sheet """ & sheetName & """ contains File URL: {2}", range.Row, range.Column, address))
			Next item

			Dim result As String = "Result-RetrieveExternalFileHyperlinks.txt"

			'Save to file.
			File.WriteAllText(result, content.ToString())

			'Launch the MS Excel file.
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
