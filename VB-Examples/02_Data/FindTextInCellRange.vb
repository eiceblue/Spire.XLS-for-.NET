Imports Spire.Xls
Imports System.Text
Imports System.IO

Namespace FindTextInCellRange
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			' Create a new Workbook object
			Dim workbook As New Workbook()

			' Load the workbook from the specified file path 
			workbook.LoadFromFile("..\..\..\..\..\..\Data\FindTextFromRangeWithFindOptions.xlsx")

			' Get the first worksheet in the workbook
			Dim sheet As Worksheet = workbook.Worksheets(0)

			' Create a StringBuilder object to store the results
			Dim builder As New StringBuilder()

			' Define the range to search for the text
			'CellRange range = sheet.Range[16, 1, 20, 2];
			Dim range As CellRange = sheet.Range("A16:B20")

			' Find all occurrences of the specified text in the range
			Dim resultRange() As CellRange = range.FindAll("e-iceblue1", FindType.Text, ExcelFindOptions.MatchEntireCellContent Or ExcelFindOptions.MatchCase)

			' Check if any occurrences were found
			If resultRange.Length <> 0 Then
				' Iterate through the found ranges and append their addresses to the StringBuilder
				For Each r As CellRange In resultRange
					Dim address As String = r.RangeAddress
					builder.AppendLine("In the range 'A16:B20', the address of the cell containing 'e-iceblue1' is: " & address)
				Next r
			End If

			' Define the output file path
			Dim result As String = "Result_out.txt"

			' Write the contents of the StringBuilder to the output file
			File.WriteAllText(result, builder.ToString())

			' Dispose the workbook object
			workbook.Dispose()

			' View the result TXT file
			OutputViewer(result)
		End Sub

		Private Sub OutputViewer(ByVal fileName As String)
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
