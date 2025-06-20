Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel
Imports Spire.Xls
Imports Spire.Xls.Core.Spreadsheet
Imports Spire.Xls.Core.Spreadsheet.PivotTables

Namespace RefreshPivotTable
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new Workbook object
            Dim workbook As New Workbook()

            ' Load an Excel file from the specified path
            workbook.LoadFromFile("..\..\..\..\..\..\Data\Template_Xls_7.xlsx")

            ' Get the second worksheet in the workbook
            Dim sheet As Worksheet = workbook.Worksheets(1)

            ' Set the value of cell D2 in the worksheet to "999"
            sheet.Range("D2").Value = "999"

            ' Get the first PivotTable in the first worksheet and cast it to XlsPivotTable
            Dim pt As XlsPivotTable = TryCast(workbook.Worksheets(0).PivotTables(0), XlsPivotTable)

            ' Enable automatic refresh of the PivotTable cache on load
            pt.Cache.IsRefreshOnLoad = True

            ' Specify the output file name
            Dim result As String = "Result-RefreshPivotTable.xlsx"

            ' Save the modified workbook to the specified file path in Excel 2013 format
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
