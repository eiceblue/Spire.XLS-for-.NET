Imports System.Collections
Imports System.ComponentModel
Imports Spire.Xls
Imports System.Text
Imports System.IO

Namespace FindStringAndNumber
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

            'Find the cells with the input string
            Dim textRanges() As CellRange = sheet.FindAllString("E-iceblue", False, False)

            'Create a string builder
            Dim builder As New StringBuilder()

            'Append the address of the found cells in string builder
            If textRanges.Length <> 0 Then
                For Each range As CellRange In textRanges
                    Dim address As String = range.RangeAddress
                    builder.AppendLine("The address of found text cell is: " & address)
                Next range
            Else
                builder.AppendLine("No cells that contain the text")
            End If

            ' Find the ells with the input integer or double data
            Dim numberRanges() As CellRange = sheet.FindAllNumber(100, True)

            'Append the address of the found cells in string builder
            If numberRanges.Length <> 0 Then
                For Each range As CellRange In numberRanges
                    Dim address As String = range.RangeAddress
                    builder.AppendLine("The address of found number cell is: " & address)
                Next range
            Else
                builder.AppendLine("No cells that contain the number")
            End If

            ' Write the contents of the string builder to a text file with the specified name.
            Dim result As String = "FindStringAndNumber_out.txt"
            File.WriteAllText(result, builder.ToString())
            ' Release the resources used by the workbook
            workbook.Dispose()

            'Launch the file
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
