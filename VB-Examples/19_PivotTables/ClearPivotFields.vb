Imports Spire.Xls
Imports System.ComponentModel
Imports System.Text
Imports Spire.Xls.Core.Spreadsheet.PivotTables

Namespace ClearPivotFields
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            'Create a new instance of Workbook
            Dim workbook As New Workbook()

            'Load the workbook from the specified file path
            workbook.LoadFromFile("..\..\..\..\..\..\Data\PivotTableExample.xlsx")

            'Get the worksheet named "PivotTable" from the workbook
            Dim sheet As Worksheet = workbook.Worksheets("PivotTable")

            'Cast the first PivotTable in the worksheet to XlsPivotTable data type
            Dim pt As XlsPivotTable = TryCast(sheet.PivotTables(0), XlsPivotTable)

            'Clear all data fields in the PivotTable
            pt.DataFields.Clear()

            'Calculate the data in the PivotTable
            pt.CalculateData()

            'Specify the filename for the resulting workbook that will be saved
            Dim result As String = "ClearPivotFields_result.xlsx"

            'Save the workbook to the specified file path in Excel 2010 format
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
