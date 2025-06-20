Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel

Imports Spire.Xls

Namespace RemoveWorksheet
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new Workbook object
            Dim workbook As New Workbook()

            ' Load an existing Excel file from the specified path
            workbook.LoadFromFile("..\..\..\..\..\..\Data\RemoveWorksheet.xlsx")

            ' Remove the worksheet at index 1 (second worksheet in the workbook)
            workbook.Worksheets.RemoveAt(1)

            ' Save the modified workbook to a file in Excel 2013 format
            workbook.SaveToFile("RemoveWorksheet_result.xlsx", ExcelVersion.Version2013)

            ' Release the resources used by the workbook
            workbook.Dispose()

            ExcelDocViewer("RemoveWorksheet_result.xlsx")
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
