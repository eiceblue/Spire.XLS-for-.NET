Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel
Imports Spire.Xls
Imports Spire.Xls.Core.Spreadsheet

Namespace ExpandAndCollapseGroups
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Creates a new instance of the Workbook class.
            Dim workbook As New Workbook()

            ' Loads an existing Excel file from the specified path.
            workbook.LoadFromFile("..\..\..\..\..\..\Data\Template_Xls_3.xlsx")

            ' Retrieves the first worksheet from the workbook.
            Dim sheet As Worksheet = workbook.Worksheets(0)

            ' Expands the grouped rows in the specified range by rows, and expands their parent groups as well.
            sheet.Range("A16:G19").ExpandGroup(GroupByType.ByRows, ExpandCollapseFlags.ExpandParent)

            ' Collapses the grouped rows in the specified range by rows.
            sheet.Range("A10:G12").CollapseGroup(GroupByType.ByRows)

            Dim result As String = "Result-ExpandAndCollapseGroups.xlsx"

            ' Saves the workbook to a file with the specified name and Excel version.
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
