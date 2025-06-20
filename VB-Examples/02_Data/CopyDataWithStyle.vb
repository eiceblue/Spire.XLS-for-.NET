Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel
Imports System.IO
Imports Spire.Xls
Imports Spire.Xls.Charts

Namespace CopyDataWithStyle
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new workbook object.
            Dim workbook As New Workbook()

            ' Get the first worksheet from the workbook.
            Dim worksheet As Worksheet = workbook.Worksheets(0)

            ' Set values for cells in the range A1 to J50 using nested loops.
            Dim cells As CellRange = worksheet.Range("A1:J50")
            For i As Integer = 1 To 10
                For j As Integer = 1 To 8
                    Dim text As String = String.Format((i - 1).ToString() & "," & (j - 1).ToString())
                    cells(i, j).Text = text
                Next j
            Next i
            ' Get a source range (A1:D3) from the worksheet.
            Dim srcRange As CellRange = worksheet.Range("A1:D3")

            ' Create a new cell style and specify font, color, and border attributes for the style.
            Dim style As CellStyle = workbook.Styles.Add("style")

            style.Font.FontName = "Calibri"
            style.Font.Color = Color.Red
            style.Borders(BordersLineType.EdgeTop).LineStyle = LineStyleType.Thin
            style.Borders(BordersLineType.EdgeTop).Color = Color.Blue
            style.Borders(BordersLineType.EdgeBottom).LineStyle = LineStyleType.Thin
            style.Borders(BordersLineType.EdgeBottom).Color = Color.Blue
            style.Borders(BordersLineType.EdgeTop).LineStyle = LineStyleType.Thin
            style.Borders(BordersLineType.EdgeTop).Color = Color.Blue
            style.Borders(BordersLineType.EdgeRight).LineStyle = LineStyleType.Thin
            style.Borders(BordersLineType.EdgeRight).Color = Color.Blue

            ' Apply the created style to the source range.
            srcRange.CellStyleName = style.Name

            ' Set the destination range in the worksheet.
            Dim destRange As CellRange = worksheet.Range("A12:D14")

            ' Copy the data from the source range to the destination range while preserving the style.
            srcRange.Copy(destRange, True, True)

            ' Define a string variable named outputFile to store the output file name as "Output.xlsx".
            Dim outputFile As String = "Output.xlsx"

            ' Save the modified workbook to a file with the name specified in outputFile, using Excel 2013 format.
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
