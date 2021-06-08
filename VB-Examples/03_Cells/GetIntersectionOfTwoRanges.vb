Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel
Imports Spire.Xls
Imports Spire.Xls.Core.Spreadsheet
Imports System.Text
Imports System.IO

Namespace GetIntersectionOfTwoRanges
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a workbook.
			Dim workbook As New Workbook()

			'Load the file from disk.
			workbook.LoadFromFile("..\..\..\..\..\..\Data\Template_Xls_1.xlsx")

			'Get the first worksheet.
			Dim sheet As Worksheet = workbook.Worksheets(0)

			'Get the two ranges.
			Dim range As CellRange = sheet.Range("A2:D7").Intersect(sheet.Range("B2:E8"))

			Dim content As New StringBuilder()
			content.AppendLine("The intersection of the two ranges ""A2:D7"" and ""B2:E8"" is:")

			'Get the intersection of the two ranges.
			For Each r As CellRange In range
				content.AppendLine(r.Value.ToString())
			Next r

			Dim result As String = "Result-GetTheIntersectionOfTwoRanges.txt"

			'Save to file.
			File.WriteAllText(result,content.ToString())

			'Launch the file.
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
