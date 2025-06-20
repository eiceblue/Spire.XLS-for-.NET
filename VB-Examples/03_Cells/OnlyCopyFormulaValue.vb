Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel
Imports Spire.Xls
Imports Spire.Xls.Core.Spreadsheet

Namespace OnlyCopyFormulaValue
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            'Creates a new instance of the Workbook class.
            Dim workbook As New Workbook()
            'Loads an Excel file from the specified path.
            workbook.LoadFromFile("..\..\..\..\..\..\Data\CopyOnlyFormulaValue1.xlsx")
            'Retrieves the first worksheet from the workbook.
            Dim sheet As Worksheet = workbook.Worksheets(0)

            'Specifies the copy option to only copy formula values.
            Dim copyOptions As CopyRangeOptions = CopyRangeOptions.OnlyCopyFormulaValue
            'Specifies the source range to be copied.
            Dim sourceRange As CellRange = sheet.Range("A6:E6")
            'Copies the source range to the destination range with the specified copy options.
            sheet.Copy(sourceRange, sheet.Range("A8:E8"), copyOptions)
            'Copies the source range to another destination range with the same copy options.
            sourceRange.Copy(sheet.Range("A10:E10"), copyOptions)
            'Specifies the name of the output file.
            Dim result As String = "Result-OnlyCopyFormulaValue.xlsx"

            'Saves the workbook to the specified file in Excel 2013 format.
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
