Imports System.Collections
Imports System.ComponentModel
Imports Spire.Xls
Imports System.Text
Imports System.IO

Namespace FindDataInSpecificRange
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a workbook
			Dim workbook As New Workbook()

			'Load the document from disk
			workbook.LoadFromFile("..\..\..\..\..\..\Data\FindCellsSample.xlsx")

			'Get the first worksheet
			Dim sheet As Worksheet = workbook.Worksheets(0)

			'Specify a range
			Dim range As CellRange = sheet.Range(1, 1, 12, 8)

			'Create a string builder
			Dim builder As New StringBuilder()

			'Find text from this range
			FindTextFromRange(range, builder)

			'Find number from this range
			FindNumberFromRange(range, builder)

			'Save to txt file
			Dim result As String = "FindDataInSpecificRange_out.txt"
			File.WriteAllText(result, builder.ToString())

			'Launch the file
			OutputViewer(result)
		End Sub
		Private Sub FindTextFromRange(ByVal range As CellRange, ByVal builder As StringBuilder)
			'Find string from this range
			Dim textRanges() As CellRange = range.FindAllString("E-iceblue", False, False)

			'Append the address of found cells in builder
			If textRanges.Length <> 0 Then
				For Each r As CellRange In textRanges
					Dim address As String = r.RangeAddress
					builder.AppendLine("The address of found text cell is: " & address)
				Next r
			Else
				builder.AppendLine("No cell contain the text")
			End If
		End Sub
		Private Sub FindNumberFromRange(ByVal range As CellRange, ByVal builder As StringBuilder)
			'Find number from this range
			Dim numberRanges() As CellRange = range.FindAllNumber(100, True)

			'Append the address of found cells in builder
			If numberRanges.Length <> 0 Then
				For Each r As CellRange In numberRanges
					Dim address As String = r.RangeAddress
					builder.AppendLine("The address of found number cell is: " & address)
				Next r
			Else
				builder.AppendLine("No cell contain the number")
			End If
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
