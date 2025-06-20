Imports Spire.Xls
Imports System.ComponentModel
Imports System.Text
Imports Spire.Xls.Core.Spreadsheet.PivotTables

Namespace UpdateDataSource
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new Workbook object
            Dim workbook As New Workbook()

            ' Load the Excel file "PivotTableExample.xlsx" from the specified path
            workbook.LoadFromFile("..\..\..\..\..\..\Data\PivotTableExample.xlsx")

            ' Get the worksheet named "Data" from the workbook
            Dim data As Worksheet = workbook.Worksheets("Data")

            ' Set the value of cell A2 in the "Data" worksheet to "NewValue"
            data.Range("A2").Text = "NewValue"

            ' Set the numeric value of cell D2 in the "Data" worksheet to 28000
            data.Range("D2").NumberValue = 28000

            ' Get the worksheet named "PivotTable" from the workbook
            Dim sheet As Worksheet = workbook.Worksheets("PivotTable")

            ' Get the first PivotTable from the "PivotTable" worksheet
            Dim pt As XlsPivotTable = TryCast(sheet.PivotTables(0), XlsPivotTable)

            ' Enable automatic refresh of the PivotTable cache on load
            pt.Cache.IsRefreshOnLoad = True

            ' Calculate the data in the PivotTable
            pt.CalculateData()

            ' Specify the name for the resulting file as "UpdateDataSource_result.xlsx"
            Dim result As String = "UpdateDataSource_result.xlsx"

            ' Save the modified workbook to the specified file in Excel 2010 format
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
