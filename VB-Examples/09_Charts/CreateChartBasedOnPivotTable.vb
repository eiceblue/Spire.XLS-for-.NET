Imports Spire.Xls
Imports Spire.Xls.Core.Spreadsheet.PivotTables
Imports System.ComponentModel
Imports System.Text

Namespace CreateChartBasedOnPivotTable
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new Workbook object
            Dim workbook As New Workbook()

            ' Load the Excel file "PivotTable.xlsx" from the specified path
            workbook.LoadFromFile("..\..\..\..\..\..\Data\PivotTable.xlsx")

            ' Get the first worksheet from the workbook
            Dim sheet As Worksheet = workbook.Worksheets(0)

            ' Retrieve the first PivotTable from the worksheet and cast it to XlsPivotTable type
            Dim pt As XlsPivotTable = TryCast(sheet.PivotTables(0), XlsPivotTable)

            ' Add a clustered bar chart to the second worksheet, based on the PivotTable data
            workbook.Worksheets(1).Charts.Add(ExcelChartType.BarClustered, pt)

            ' Specify the output file name as "CreateChartBasedOnPivotTable.xlsx"
            Dim output As String = "CreateChartBasedOnPivotTable.xlsx"

            ' Save the modified workbook to the specified file path, using Excel 2013 format
            workbook.SaveToFile(output, ExcelVersion.Version2013)
            ' Release the resources used by the workbook
            workbook.Dispose()

            'View the document
            FileViewer(output)
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
