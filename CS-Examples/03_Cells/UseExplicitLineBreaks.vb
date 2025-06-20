Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel
Imports System.IO
Imports Spire.Xls
Imports Spire.Xls.Charts

Namespace UseExplicitLineBreaks
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            'Creates a new instance of the Workbook class.
            Dim workbook As New Workbook()

            'Retrieves the first worksheet from the workbook.
            Dim sheet1 As Worksheet = workbook.Worksheets(0)

            'Specifies the cell range as "C5".
            Dim c5 As CellRange = sheet1.Range("C5")

            'Sets the width of the column containing the cell range to 70.
            sheet1.SetColumnWidth(c5.Column, 70)

            'Put the string value with explicit line breaks
            c5.Value = "Spire.XLS is a professional Excel API" & vbLf & " that can be used to create, read, " & vbLf & "write, convert and print Excel files in any type " & vbLf & "of application. " & vbLf & "Spire.XLS offers object model" & vbLf & " Excel API for speeding up Excel programming in .NET/Java/C++/Python platform -" & vbLf & " create new Excel documents from template, edit existing " & vbLf & "Excel documents and " & vbLf & "convert Excel files."

            'Enables the text wrap feature for the cell range, causing the text to wrap within the cell.
            c5.IsWrapText = True

            'Specifies the name of the output file.
            Dim outputFile As String = "Output.xlsx"

            'Saves the workbook to the specified file in Excel 2013 format.
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
