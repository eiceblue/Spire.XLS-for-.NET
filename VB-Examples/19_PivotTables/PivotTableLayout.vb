Imports Spire.Xls
Imports Spire.Xls.Core.Spreadsheet.PivotTables
Imports System.ComponentModel
Imports System.Text

Namespace PivotTableLayout
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new Workbook object
            Dim workbook As New Workbook()

            ' Load the Excel file from a specified path
            workbook.LoadFromFile("..\..\..\..\..\..\Data\PivotTable.xlsx")

            ' Get the first worksheet in the workbook
            Dim worksheet As Worksheet = workbook.Worksheets(0)

            ' Retrieve the PivotTable object from the worksheet and cast it as an XlsPivotTable
            Dim xlsPivotTable As XlsPivotTable = CType(worksheet.PivotTables(0), XlsPivotTable)

            ' Set the report layout of the PivotTable to Tabular
            xlsPivotTable.Options.ReportLayout = PivotTableLayoutType.Tabular

            ' Specify the file name for the result file
            Dim result As String = "PivotLayoutTabular_output.xlsx"

            ' Save the workbook to the specified file path with the specified Excel version (2013)
            workbook.SaveToFile(result, ExcelVersion.Version2013)

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

	End Class
End Namespace
