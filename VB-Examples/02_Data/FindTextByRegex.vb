Imports Spire.Xls
Imports System.IO

Namespace FindTextByRegex
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			' Load an existing workbook from a file
			Dim workbook As New Workbook()
			workbook.LoadFromFile("..\..\..\..\..\..\Data\FindTextByRegex.xlsx")

			' Get the first sheet
			Dim worksheet As Worksheet = workbook.Worksheets(0)

			' Find cell ranges by Regex
			Dim ranges() As CellRange = worksheet.FindAllString(".*North.", False, False, True)
			Dim information As String = ""

			' Get the information of every cell range
			For Each range As CellRange In ranges
				information &= "RangeAddressLocal:" & range.RangeAddressLocal & vbCrLf
				information &= "Text:" & range.Text & vbCrLf
			Next range

			' Specify the output file name for the result
			Dim result As String = "FindTextByRegex_result.txt"

			File.WriteAllText(result, information)

			' Dispose of the workbook object to release resources
			workbook.Dispose()

			' Launch the file
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
