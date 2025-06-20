Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel
Imports Spire.Xls
Imports Spire.Xls.Core.Spreadsheet

Namespace UnmergeCells
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            'Creates a new instance of the Workbook class.
            Dim workbook As New Workbook()

            'Loads an Excel file from the specified path.
            workbook.LoadFromFile("..\..\..\..\..\..\Data\Template_Xls_1.xlsx")

            'Retrieves the first worksheet from the workbook.
            Dim sheet As Worksheet = workbook.Worksheets(0)

            'Unmerges the cells in range "F2".
            sheet.Range("F2").UnMerge()

            'Unmerges the cells in range "F7".
            sheet.Range("F7").UnMerge()
            'Specifies the name of the output file.
            Dim result As String = "Result-UnmergeCells.xlsx"

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
