Imports Spire.Xls
Imports System.ComponentModel
Imports System.Text
Imports Spire.Xls.Core.Spreadsheet.PivotTables

Namespace FormattingAppearance
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new Workbook object
            Dim workbook As New Workbook()

            ' Load an Excel file from the specified path
            workbook.LoadFromFile("..\..\..\..\..\..\Data\PivotTableExample.xlsx")

            ' Get the worksheet with the name "PivotTable"
            Dim sheet As Worksheet = workbook.Worksheets("PivotTable")

            ' Get the first PivotTable in the worksheet and cast it to XlsPivotTable
            Dim pt As XlsPivotTable = TryCast(sheet.PivotTables(0), XlsPivotTable)

            ' Set the built-in style of the PivotTable to PivotStyleLight10
            pt.BuiltInStyle = PivotBuiltInStyles.PivotStyleLight10

            ' Enable the display of grid drop zone in the PivotTable
            pt.Options.ShowGridDropZone = True

            ' Set the row layout of the PivotTable to Compact
            pt.Options.RowLayout = PivotTableLayoutType.Compact

            ' Specify the output file name
            Dim result As String = "FormattingAppearance_result.xlsx"

            ' Save the modified workbook to the specified file path in Excel 2010 format
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
