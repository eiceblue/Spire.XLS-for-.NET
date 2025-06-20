Imports Spire.Xls
Imports System.ComponentModel
Imports System.Text
Imports Spire.Xls.Core.Spreadsheet.PivotTables

Namespace ConsolidationFunctions
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new Workbook object
            Dim workbook As New Workbook()

            ' Load an Excel file named "PivotTableExample.xlsx" from a specific location
            workbook.LoadFromFile("..\..\..\..\..\..\Data\PivotTableExample.xlsx")

            ' Get the worksheet named "PivotTable"
            Dim sheet As Worksheet = workbook.Worksheets("PivotTable")

            ' Get the first PivotTable on the sheet and cast it to XlsPivotTable type
            Dim pt As XlsPivotTable = TryCast(sheet.PivotTables(0), XlsPivotTable)

            ' Set the subtotal type for the first data field as Average
            pt.DataFields(0).Subtotal = SubtotalTypes.Average

            ' Set the subtotal type for the second data field as Maximum
            pt.DataFields(1).Subtotal = SubtotalTypes.Max

            ' Calculate the data of the PivotTable
            pt.CalculateData()

            ' Specify the filename for the resulting Excel file
            Dim result As String = "ConsolidationFunctions_result.xlsx"

            ' Save the Workbook object to the specified file in Excel 2010 format
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
