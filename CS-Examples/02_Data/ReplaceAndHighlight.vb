Imports Spire.Xls
Imports System.ComponentModel
Imports System.Text

Namespace ReplaceAndHighlight
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new Workbook object
            Dim workbook As New Workbook()

            ' Load the Excel file from the specified path
            workbook.LoadFromFile("..\..\..\..\..\..\Data\ReplaceAndHighlight.xlsx")

            ' Get the first worksheet from the workbook
            Dim worksheet As Worksheet = workbook.Worksheets(0)

            ' Find all occurrences of the string "Total" in the worksheet, case-sensitive and whole word matching
            Dim ranges() As CellRange = worksheet.FindAllString("Total", True, True)

            ' Iterate through each found range
            For Each range As CellRange In ranges
                ' Replace the text with "Sum"
                range.Text = "Sum"

                ' Set the color of the range to yellow
                range.Style.Color = Color.Yellow
            Next range

            ' Specify the output file name for the modified workbook
            Dim result As String = "ReplaceAndHighlight_result.xlsx"

            ' Save the workbook to the specified file path using Excel 2010 format
            workbook.SaveToFile(result, ExcelVersion.Version2010)
            ' Release the resources used by the workbook
            workbook.Dispose()

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
