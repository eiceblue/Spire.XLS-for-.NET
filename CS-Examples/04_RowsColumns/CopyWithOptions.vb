Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel
Imports System.IO
Imports Spire.Xls
Imports Spire.Xls.Charts

Namespace CopyWithOptions
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            'Instantiate a new Workbook object to create a workbook.
            Dim workbook As New Workbook()

            'Load the Excel document from the specified path.
            workbook.LoadFromFile("..\..\..\..\..\..\Data\Sample.xlsx")

            'Get the first worksheet from the workbook.
            Dim sheet1 As Worksheet = workbook.Worksheets(0)

            'Add a new worksheet named "DestSheet" to the workbook.
            Dim destinationSheet As Worksheet = workbook.Worksheets.Add("DestSheet")

            'Specify the range of cells (B2:D4) that will be copied.
            Dim cellRange As CellRange = sheet1.Range("B2:D4")

            'Copy the content of the specified range from the first worksheet to the second worksheet, starting at cell C3 (row offset: 2, column offset: 1).
            'Maintain the original styles of the copied cells and update any cell references in formulas accordingly.
            workbook.Worksheets(0).Copy(cellRange, workbook.Worksheets(1), 2, 1, True, True)


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
