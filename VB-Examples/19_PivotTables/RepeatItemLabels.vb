Imports Spire.Xls
Imports System.ComponentModel
Imports System.Text
Imports Spire.Xls.Core.Spreadsheet.PivotTables

Namespace RepeatItemLabels
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new Workbook object
            Dim workbook As New Workbook()

            ' Load an Excel file from the specified path
            workbook.LoadFromFile("..\..\..\..\..\..\Data\RepeatItemLabelsExample.xlsx")

            ' Get the first worksheet in the workbook
            Dim sheet As Worksheet = workbook.Worksheets(0)

            ' Create a new empty worksheet
            Dim sheet2 As Worksheet = workbook.CreateEmptySheet()
            sheet2.Name = "Pivot Table"

            ' Define the range of data to be used for the PivotTable
            Dim dataRange As CellRange = sheet.Range("A1:D9")

            ' Create a PivotCache based on the data range
            Dim cache As PivotCache = workbook.PivotCaches.Add(dataRange)

            ' Add a PivotTable to sheet2 with the specified name, source range, and cache
            Dim pt As PivotTable = sheet2.PivotTables.Add("Pivot Table", sheet.Range("A1"), cache)

            ' Configure the first field ("VendorNo") in the PivotTable
            Dim r1 = pt.PivotFields("VendorNo")
            r1.Axis = AxisTypes.Row
            pt.Options.RowHeaderCaption = "VendorNo"
            r1.Subtotals = SubtotalTypes.None
            r1.RepeatItemLabels = True

            ' Configure the second field ("OnHand") in the PivotTable
            pt.PivotFields("OnHand").RepeatItemLabels = True

            ' Set the row layout of the PivotTable to Tabular
            pt.Options.RowLayout = PivotTableLayoutType.Tabular

            ' Configure the third field ("Desc") in the PivotTable
            Dim r2 = pt.PivotFields("Desc")
            r2.Axis = AxisTypes.Row

            ' Add a data field ("OnHand") to the PivotTable
            pt.DataFields.Add(pt.PivotFields("OnHand"), "Sum of onHand", SubtotalTypes.None)

            ' Set the built-in style of the PivotTable to PivotStyleMedium12
            pt.BuiltInStyle = PivotBuiltInStyles.PivotStyleMedium12

            ' Specify the output file name
            Dim result As String = "RepeatItemLabels_result.xlsx"

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

		Private Sub Form1_Load(ByVal sender As Object, ByVal e As EventArgs) Handles MyBase.Load

		End Sub
	End Class
End Namespace
