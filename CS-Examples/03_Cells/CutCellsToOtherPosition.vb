Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel

Imports Spire.Xls

Namespace CutCellsToOtherPosition
	Partial Public Class Form1
		Inherits Form

		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            'Instantiate a new workbook object.
            Dim workbook As New Workbook()

            'Load an existing Excel document from the specified file path.
            workbook.LoadFromFile("..\..\..\..\..\..\Data\SampleB_2.xlsx")
            'Retrieve the first worksheet from the workbook.
            Dim sheet As Worksheet = workbook.Worksheets(0)
            'Define a cell range for the original cells.
            Dim Ori As CellRange = sheet.Range("A1:C5")
            'Define a cell range for the destination cells.
            Dim Dest As CellRange = sheet.Range("A26:C30")

            'Copy the content and formatting of the original cells to the destination cells, including values, formulas, and styles.
            sheet.Copy(Ori, Dest, True, True, True)

            'Iterate through each cell range within the original range.
            For Each cr As CellRange In Ori
                'Clear all content, including values, formulas, and formatting, in the current cell range.
                cr.ClearAll()
            Next cr

            'Specify the file name for the output file.
            Dim result As String = "result.xlsx"
            'Save the modified workbook to the specified output file using Excel 2010 version.
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
