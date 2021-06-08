Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel
Imports Spire.Xls
Imports Spire.Xls.Core.Spreadsheet

Namespace RetrieveAndExtractData
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a new workbook instance and get the first worksheet.
			Dim newBook As New Workbook()
			Dim newSheet As Worksheet = newBook.Worksheets(0)

			'Create a new workbook instance and load the sample Excel file.
			Dim workbook As New Workbook()
			workbook.LoadFromFile("..\..\..\..\..\..\Data\Template_Xls_3.xlsx")

			'Get the first worksheet.
			Dim sheet As Worksheet = workbook.Worksheets(0)

			'Retrieve data and extract it to the first worksheet of the new excel workbook.
			Dim i As Integer = 1
			Dim columnCount As Integer = sheet.Columns.Length
			For Each range As CellRange In sheet.Columns(0)
				If range.Text = "teacher" Then
					Dim sourceRange As CellRange = sheet.Range(range.Row, 1, range.Row, columnCount)
					Dim destRange As CellRange = newSheet.Range(i, 1, i, columnCount)
					sheet.Copy(sourceRange, destRange,True)
					i += 1
				End If
			Next range

			Dim result As String = "Result-RetrieveAndExtractDataToNewExcelFile.xlsx"

			'Save to file.
			newBook.SaveToFile(result, ExcelVersion.Version2013)

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
