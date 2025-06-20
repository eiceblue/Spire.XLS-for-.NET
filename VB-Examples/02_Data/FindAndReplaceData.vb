Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel
Imports System.IO
Imports Spire.Xls
Imports Spire.Xls.Charts
Imports System.Text

Namespace FindAndReplaceData
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Instantiate a new workbook object to store the Excel file.
            Dim workbook As New Workbook()

            ' Load the Excel file from the specified path.
            workbook.LoadFromFile("..\..\..\..\..\..\Data\CreateTable.xlsx")

            'Access the first worksheet in the workbook.
            Dim worksheet As Worksheet = workbook.Worksheets(0)

            ' Find all occurrences of the string "Area" in the worksheet, ignoring case sensitivity and exact match.
            Dim ranges() As CellRange = worksheet.FindAllString("Area", False, False)

            ' Iterate through each found range.
            For Each range As CellRange In ranges
                ' Replace the text in the range with "Area Code".
                range.Text = "Area Code"
                ' Set the background color of the range to yellow.
                range.Style.Color = Color.Yellow
            Next range

            ' Specify the file name for the output Excel file.
            Dim outputFile As String = "Output.xlsx"

            ' Save the modified workbook to a new Excel file with the specified name and version.
            workbook.SaveToFile(outputFile, ExcelVersion.Version2013)

            ' Release the resources used by the workbook
            workbook.Dispose()

            'Launching the output file.
            Viewer(outputFile)
		End Sub
		Private Sub Viewer(ByVal fileName As String)
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
