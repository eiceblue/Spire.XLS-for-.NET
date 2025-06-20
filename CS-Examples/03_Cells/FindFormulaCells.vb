Imports System.Collections
Imports System.ComponentModel

Imports Spire.Xls
Imports System.Text
Imports System.IO

Namespace FindFormulaCells
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            'Instantiate a new workbook object.
            Dim workbook As New Workbook()

            'Load an existing Excel document from the specified file path.
            workbook.LoadFromFile("..\..\..\..\..\..\Data\FindCellsSample.xlsx")

            'Retrieve the first worksheet from the workbook.
            Dim sheet As Worksheet = workbook.Worksheets(0)

            'Search for all cell ranges in the worksheet that contain the formula "=SUM(A11,A12)".
            Dim ranges() As CellRange = sheet.FindAll("=SUM(A11,A12)", FindType.Formula, ExcelFindOptions.None)

            'Instantiate a new StringBuilder object to build the output text.
            Dim builder As New StringBuilder()

            'Check if any cell ranges containing the formula were found.
            If ranges.Length <> 0 Then
                'Iterate through each found cell range.
                For Each range As CellRange In ranges
                    'Retrieve the address of the current cell range.
                    Dim address As String = range.RangeAddress
                    'Append the address of the current cell range to the string builder.
                    builder.AppendLine("The address of found cell is: " & address)
                Next range
            Else
                'Append a message indicating that no cell contains the specified formula.
                builder.AppendLine("No cell contain the formula")
            End If

            'Specify the file name for the output file.
            Dim result As String = "FindFormulaCells_out.txt"
            'Write the contents of the string builder to the specified output file.
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
