Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel

Imports Spire.Xls
Imports Spire.Xls.Charts

Namespace FreezePanes
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new instance of Workbook class
            Dim workbook As New Workbook()
            ' Load the Excel file from the specified path
            workbook.LoadFromFile("..\..\..\..\..\..\Data\FreezePanes.xlsx")

            ' Get the first worksheet in the workbook
            Dim sheet As Worksheet = workbook.Worksheets(0)

            ' Freeze the panes at row 2 and column 1
            sheet.FreezePanes(2, 1)

            ' Save the modified workbook to a new file named "Output.xlsx" in Excel 2010 format
            workbook.SaveToFile("Output.xlsx", ExcelVersion.Version2010)

            ' Release the resources used by the workbook
            workbook.Dispose()
            ExcelDocViewer("Output.xlsx")
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
