Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel
Imports Spire.Xls
Imports Spire.Xls.Core.Spreadsheet

Namespace CopySheetWithinWorkbook
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new Workbook object
            Dim workbook As New Workbook()

            ' Load an existing Excel file into the workbook from a specified path
            workbook.LoadFromFile("..\..\..\..\..\..\Data\Template_Xls_4.xlsx")

            ' Get the reference to the first worksheet in the workbook
            Dim sheet As Worksheet = workbook.Worksheets(0)

            ' Add a new worksheet named "MySheet" to the workbook and get its reference
            Dim sheet1 As Worksheet = workbook.Worksheets.Add("MySheet")

            ' Get the range of cells that are allocated (used) in the source worksheet
            Dim sourceRange As CellRange = sheet.AllocatedRange

            ' Copy the source range of cells to the destination worksheet starting from the first row and column,
            ' while preserving the formatting and formulas
            sheet.Copy(sourceRange, sheet1, sheet.FirstRow, sheet.FirstColumn, True)

            ' Specify the name for the resulting copied sheet within the workbook
            Dim result As String = "Result-CopySheetWithinWorkbook.xlsx"

            ' Save the modified workbook to a new file with the specified name and Excel version (2013)
            workbook.SaveToFile(result, ExcelVersion.Version2013)

            ' Release the resources used by the workbook
            workbook.Dispose()

            'Launch the MS Excel file.
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
