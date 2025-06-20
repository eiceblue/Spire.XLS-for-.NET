Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel
Imports Spire.Xls
Imports Spire.Xls.Core.Spreadsheet

Namespace ExpandOrCollapseRows
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new Workbook object
            Dim workbook As New Workbook()

            ' Load the Excel file from a specified path
            workbook.LoadFromFile("..\..\..\..\..\..\Data\Template_Xls_7.xlsx")

            ' Get the first worksheet in the workbook
            Dim sheet As Worksheet = workbook.Worksheets(0)

            ' Retrieve the PivotTable object from the worksheet and cast it as an XlsPivotTable
            Dim pivotTable As Spire.Xls.Core.Spreadsheet.PivotTables.XlsPivotTable = TryCast(sheet.PivotTables(0), Spire.Xls.Core.Spreadsheet.PivotTables.XlsPivotTable)

            ' Calculate the data in the PivotTable
            pivotTable.CalculateData()

            ' Hide the detail for a specific item in the "Vendor No" field of the PivotTable
            TryCast(pivotTable.PivotFields("Vendor No"), Spire.Xls.Core.Spreadsheet.PivotTables.XlsPivotField).HideItemDetail("3501", True)

            ' Show the detail for another specific item in the "Vendor No" field of the PivotTable
            TryCast(pivotTable.PivotFields("Vendor No"), Spire.Xls.Core.Spreadsheet.PivotTables.XlsPivotField).HideItemDetail("3502", False)

            ' Specify the file name for the result file
            Dim result As String = "Result-ExpandOrCollapseRowsInPivotTable.xlsx"

            ' Save the workbook to the specified file path with the specified Excel version (2013)
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
