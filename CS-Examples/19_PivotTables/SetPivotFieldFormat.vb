Imports Spire.Xls
Imports System.ComponentModel
Imports System.Text
Imports Spire.Xls.Core.Spreadsheet.PivotTables

Namespace SetPivotFieldFormat
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new Workbook object
            Dim workbook As New Workbook()

            ' Load the workbook from a file
            workbook.LoadFromFile("..\..\..\..\..\..\Data\PivotTableExample.xlsx")

            ' Get the worksheet named "PivotTable"
            Dim sheet As Worksheet = workbook.Worksheets("PivotTable")

            ' Access the first PivotTable in the worksheet
            Dim pt As XlsPivotTable = TryCast(sheet.PivotTables(0), XlsPivotTable)

            ' Access the first PivotField in the PivotTable
            Dim pf As PivotField = TryCast(pt.PivotFields(0), PivotField)

            ' Set the sort type of the PivotField to Ascending
            pf.SortType = PivotFieldSortType.Ascending

            ' Enable displaying the top subtotal for the PivotField
            pf.SubtotalTop = True

            ' Set the subtotal type for the PivotField to Count
            pf.Subtotals = SubtotalTypes.Count

            ' Enable automatic display of the PivotField
            pf.IsAutoShow = True

            ' Save the modified workbook to a new file named "SetPivotFieldFormat_result.xlsx"
            Dim result As String = "SetPivotFieldFormat_result.xlsx"
            workbook.SaveToFile(result, ExcelVersion.Version2010)

            ' Release the resources used by the workbook
            workbook.Dispose()

            'View the document
            FileViewer(result)
		End Sub

		Private Sub FileViewer(ByVal fileName As String)
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
