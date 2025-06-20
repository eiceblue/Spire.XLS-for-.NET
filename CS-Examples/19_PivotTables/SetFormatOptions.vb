Imports Spire.Xls
Imports System.ComponentModel
Imports System.Text
Imports Spire.Xls.Core.Spreadsheet.PivotTables

Namespace SetFormatOptions
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new instance of Workbook class
            Dim workbook As New Workbook()

            ' Load the Excel file into the Workbook object
            workbook.LoadFromFile("..\..\..\..\..\..\Data\PivotTableExample.xlsx")

            ' Access the specific worksheet by name and assign it to a Worksheet object
            Dim sheet As Worksheet = workbook.Worksheets("PivotTable")

            ' Retrieve the first PivotTable on the worksheet and cast it to XlsPivotTable object
            Dim pt As XlsPivotTable = TryCast(sheet.PivotTables(0), XlsPivotTable)

            ' Enable automatic formatting for the PivotTable
            pt.Options.IsAutoFormat = True

            ' Show grand totals for rows in the PivotTable
            pt.ShowRowGrand = True

            ' Show grand totals for columns in the PivotTable
            pt.ShowColumnGrand = True

            ' Set the option to display a specified string for null values in the PivotTable
            pt.DisplayNullString = True
            pt.NullString = "null"

            ' Set the page field order for multiple page fields in the PivotTable
            pt.PageFieldOrder = PagesOrderType.DownThenOver

            ' Specify the output file name after saving the modified Workbook
            Dim result As String = "SetFormatOptions_result.xlsx"
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
