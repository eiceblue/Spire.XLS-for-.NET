Imports Spire.Xls
Imports System.ComponentModel
Imports System.Text

Namespace SortPivotTable
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new instance of Workbook
            Dim workbook As New Workbook()

            ' Load an existing Excel file from the specified path
            workbook.LoadFromFile("..\..\..\..\..\..\Data\SortPivotTable.xlsx")

            ' Get the first worksheet from the workbook
            Dim sheet As Worksheet = workbook.Worksheets(0)

            ' Create a new empty worksheet in the workbook
            Dim sheet2 As Worksheet = workbook.CreateEmptySheet()

            ' Set the name of the new worksheet to "Pivot Table"
            sheet2.Name = "Pivot Table"

            ' Define a range of cells as data range for the PivotTable
            Dim dataRange As CellRange = sheet.Range("A1:C9")

            ' Create a PivotCache using the data range
            Dim cache As PivotCache = workbook.PivotCaches.Add(dataRange)

            ' Add a PivotTable to the new worksheet, using the PivotCache and specifying the starting cell
            Dim pt As PivotTable = sheet2.PivotTables.Add("Pivot Table", sheet.Range("A1"), cache)

            ' Get the PivotField for the "No" field and set its axis type to Row
            Dim r1 As PivotField = TryCast(pt.PivotFields("No"), PivotField)
            r1.Axis = AxisTypes.Row

            ' Set the layout type of the PivotTable to Tabular
            pt.Options.RowLayout = PivotTableLayoutType.Tabular

            ' Set the sort type of the "No" field to Descending
            r1.SortType = PivotFieldSortType.Descending

            ' Get the PivotField for the "Name" field and set its axis type to Row
            Dim r2 As PivotField = TryCast(pt.PivotFields("Name"), PivotField)
            r2.Axis = AxisTypes.Row

            ' Add a data field to the PivotTable, using the "OnHand" field and specifying the aggregation function and subtotal type
            pt.DataFields.Add(pt.PivotFields("OnHand"), "Sum of onHand", SubtotalTypes.None)

            ' Set the built-in style of the PivotTable to PivotStyleMedium12
            pt.BuiltInStyle = PivotBuiltInStyles.PivotStyleMedium12

            ' Save the modified workbook to a new file with the specified name and Excel version
            Dim result As String = "SortPivotTable_result.xlsx"
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

		Private Sub btnClose_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnClose.Click
			Close()
		End Sub
	End Class
End Namespace
