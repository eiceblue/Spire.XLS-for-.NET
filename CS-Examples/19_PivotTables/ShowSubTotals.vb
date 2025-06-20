Imports Spire.Xls
Imports System.ComponentModel
Imports System.Text
Imports Spire.Xls.Core.Spreadsheet.PivotTables

Namespace ShowSubTotals
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new instance of Workbook
            Dim workbook As New Workbook()

            ' Load an existing Excel file from the specified path
            workbook.LoadFromFile("..\..\..\..\..\..\Data\ShowSubTotals.xlsx")

            ' Get the worksheet named "Pivot Table"
            Dim sheet As Worksheet = workbook.Worksheets("Pivot Table")

            ' Get the first PivotTable from the worksheet and cast it to XlsPivotTable
            Dim pt As XlsPivotTable = TryCast(sheet.PivotTables(0), XlsPivotTable)

            ' Enable showing subtotals in the PivotTable
            pt.ShowSubtotals = True

            ' Specify the name of the resulting file
            Dim result As String = "ShowSubTotals_result.xlsx"

            ' Save the modified workbook to a new file with the specified name and Excel version (2010)
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
