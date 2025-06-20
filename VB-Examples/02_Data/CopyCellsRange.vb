Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel
Imports System.IO
Imports Spire.Xls
Imports Spire.Xls.Charts

Namespace CopyCellsRange
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new workbook object.
            Dim workbook As New Workbook()

            ' Load an Excel document named "CreateTable.xlsx" from the specified file path.
            workbook.LoadFromFile("..\..\..\..\..\..\Data\CreateTable.xlsx")

            ' Get the first worksheet from the workbook.
            Dim sheet1 As Worksheet = workbook.Worksheets(0)

            ' Specify a destination range in the worksheet (cells G1 to H19).
            Dim cells As CellRange = sheet1.Range("G1:H19")

            ' Copy the selected range (cells B1 to C19) to the destination range specified.
            sheet1.Range("B1:C19").Copy(cells)


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
