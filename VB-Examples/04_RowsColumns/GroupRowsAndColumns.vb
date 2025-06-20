Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel

Imports Spire.Xls
Imports Spire.Xls.Charts

Namespace GroupRowsAndColumns
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new Workbook object.
            Dim workbook As New Workbook()
            ' Load an existing Excel file into the Workbook object.
            workbook.LoadFromFile("..\..\..\..\..\..\Data\GroupRowsAndColumns.xls")
            ' Get the first worksheet from the Workbook.
            Dim sheet As Worksheet = workbook.Worksheets(0)

            ' Group rows from 1 to 5 without collapsing the group.
            sheet.GroupByRows(1, 5, False)

            ' Group columns from 1 to 3 without collapsing the group.
            sheet.GroupByColumns(1, 3, False)
            ' Save the modified Workbook to a new file.
            workbook.SaveToFile("GroupRowsAndColumns.xlsx", ExcelVersion.Version2010)
            ' Release the resources used by the workbook
            workbook.Dispose()

            ExcelDocViewer("GroupRowsAndColumns.xlsx")
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
