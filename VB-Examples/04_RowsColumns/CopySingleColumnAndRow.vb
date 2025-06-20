Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel
Imports System.IO
Imports Spire.Xls
Imports Spire.Xls.Charts

Namespace CopySingleColumnAndRow
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            'Instantiate a new Workbook object to create a workbook.
            Dim workbook As New Workbook()

            'Load the Excel document from the specified path.
            workbook.LoadFromFile("..\..\..\..\..\..\Data\CreateTable.xlsx")

            'Get the first worksheet from the workbook.
            Dim sheet1 As Worksheet = workbook.Worksheets(0)

            'Specify the destination range where the copied column will be placed.n 
            Dim columnCells As CellRange = sheet1.Range("G1:G19")

            'Copy the content of the second column and paste it into the destination range.
            sheet1.Columns(1).Copy(columnCells)

            'Specify the destination range where the copied row will be placed.
            Dim rowCells As CellRange = sheet1.Range("A21:E21")

            'Copy the content of the first row and paste it into the destination range.
            sheet1.Rows(0).Copy(rowCells)

            'Specify the name for the resulting file.
            Dim outputFile As String = "Output.xlsx"

            'Save the workbook to a file with the specified name and Excel version (in this case, Excel 2013).
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
