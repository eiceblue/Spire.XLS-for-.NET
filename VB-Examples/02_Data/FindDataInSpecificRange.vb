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
            ' Instantiate a new workbook object.
            Dim workbook As New Workbook()

            ' Load the Excel file from the specified path.
            workbook.LoadFromFile("..\..\..\..\..\..\Data\FindCellsSample.xlsx")

            ' Access the first worksheet in the workbook.
            Dim sheet As Worksheet = workbook.Worksheets(0)

            ' Define a range of cells from row 1, column 1 to row 12, column 8.
            Dim range As CellRange = sheet.Range(1, 1, 12, 8)

            ' Instantiate a new StringBuilder object to store the found data.
            Dim builder As New StringBuilder()

            ' Search for text data within the specified range and append the results to the string builder.
            FindTextFromRange(range, builder)

            ' Search for numerical data within the specified range and append the results to the string builder.
            FindNumberFromRange(range, builder)

            ' Specify the file name for the output text file.
            Dim result As String = "FindDataInSpecificRange_out.txt"

            ' Write the contents of the string builder to a text file with the specified name.
            File.WriteAllText(result, builder.ToString())
            ' Release the resources used by the workbook
            workbook.Dispose()

            'Launch the file
            OutputViewer(result)
		End Sub
		Private Sub FindTextFromRange(ByVal range As CellRange, ByVal builder As StringBuilder)
            ' Find all occurrences of the string "E-iceblue" within the specified range, ignoring case sensitivity and exact match.
            Dim textRanges() As CellRange = range.FindAllString("E-iceblue", False, False)

            ' Check if any text ranges were found.
            If textRanges.Length <> 0 Then
                ' Iterate through each found text range.
                For Each r As CellRange In textRanges
                    ' Get the address of the current text cell range.
                    Dim address As String = r.RangeAddress
                    ' Append the address of the found text cell to the string builder.
                    builder.AppendLine("The address of found text cell is: " & address)
                Next r
            Else
                ' If no text ranges were found, append a message to the string builder indicating that no cell contains the specified text.
                builder.AppendLine("No cell contains the text")
            End If
        End Sub
		Private Sub FindNumberFromRange(ByVal range As CellRange, ByVal builder As StringBuilder)
            ' Find all numbers greater than or equal to 100 within the specified range.
            Dim numberRanges() As CellRange = range.FindAllNumber(100, True)

            ' Check if any number ranges were found.
            If numberRanges.Length <> 0 Then
                ' Iterate through each found number range.
                For Each r As CellRange In numberRanges
                    ' Get the address of the current number cell range.
                    Dim address As String = r.RangeAddress
                    ' Append the address of the found number cell to the string builder.
                    builder.AppendLine("The address of found number cell is: " & address)
                Next r
            Else
                ' If no number ranges were found, append a message to the string builder indicating that no cell contains a number.
                builder.AppendLine("No cell contains the number")
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
