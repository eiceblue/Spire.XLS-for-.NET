Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel

Imports Spire.Xls

Namespace CreatePivotTable
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new Workbook object
            Dim workbook As New Workbook()

            ' Get the first worksheet of the workbook
            Dim sheet As Worksheet = workbook.Worksheets(0)

            ' Set values for the headers in the worksheet
            sheet.Range("A1").Value = "Product"
            sheet.Range("B1").Value = "Month"
            sheet.Range("C1").Value = "Count"

            ' Enter data into the worksheet
            sheet.Range("A2").Value = "SpireDoc"
            sheet.Range("A3").Value = "SpireDoc"
            sheet.Range("A4").Value = "SpireXls"
            sheet.Range("A5").Value = "SpireDoc"
            sheet.Range("A6").Value = "SpireXls"
            sheet.Range("A7").Value = "SpireXls"

            sheet.Range("B2").Value = "January"
            sheet.Range("B3").Value = "February"
            sheet.Range("B4").Value = "January"
            sheet.Range("B5").Value = "January"
            sheet.Range("B6").Value = "February"
            sheet.Range("B7").Value = "February"

            sheet.Range("C2").Value = "10"
            sheet.Range("C3").Value = "15"
            sheet.Range("C4").Value = "9"
            sheet.Range("C5").Value = "7"
            sheet.Range("C6").Value = "8"
            sheet.Range("C7").Value = "10"

            ' Define a range for the data in the worksheet
            Dim dataRange As CellRange = sheet.Range("A1:C7")

            ' Create a PivotCache using the data range
            Dim cache As PivotCache = workbook.PivotCaches.Add(dataRange)

            ' Add a PivotTable to the worksheet using the PivotCache
            Dim pt As PivotTable = sheet.PivotTables.Add("Pivot Table", sheet.Range("E10"), cache)

            ' Set the "Product" field as a row axis in the PivotTable
            Dim pf As PivotField = TryCast(pt.PivotFields("Product"), PivotField)
            pf.Axis = AxisTypes.Row

            ' Set the "Month" field as another row axis in the PivotTable
            Dim pf2 As PivotField = TryCast(pt.PivotFields("Month"), PivotField)
            pf2.Axis = AxisTypes.Row

            ' Add the "Count" field to the PivotTable as a data field, with the name "SUM of Count"
            pt.DataFields.Add(pt.PivotFields("Count"), "SUM of Count", SubtotalTypes.Sum)

            ' Set the PivotTable style to PivotStyleMedium12
            pt.BuiltInStyle = PivotBuiltInStyles.PivotStyleMedium12

            ' Calculate the data in the PivotTable
            pt.CalculateData()

            ' Auto-fit columns 5 and 6 in the worksheet for better display of the PivotTable
            sheet.AutoFitColumn(5)
            sheet.AutoFitColumn(6)

            ' Save the workbook to a file named "CreatePivotTable_output.xlsx" in Excel 2010 format
            Dim result As String = "CreatePivotTable_output.xlsx"
            workbook.SaveToFile(result, ExcelVersion.Version2010)

            ' Release the resources used by the workbook
            workbook.Dispose()

            ExcelDocViewer(result)
		End Sub

		Private Sub ExcelDocViewer(ByVal fileName As String)
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
