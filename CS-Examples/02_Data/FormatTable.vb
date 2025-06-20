Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel

Imports Spire.Xls
Imports Spire.Xls.Core.Spreadsheet.PivotTables

Namespace FormatTable
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new Workbook object
            Dim workbook As New Workbook()
            ' Load an existing Excel file 
            workbook.LoadFromFile("..\..\..\..\..\..\Data\FormatTable.xlsx")
            ' Get the first worksheet from the workbook
            Dim sheet As Worksheet = workbook.Worksheets(0)

            ' Add Default Style to the table
            sheet.ListObjects(0).BuiltInTableStyle = TableBuiltInStyles.TableStyleMedium9

            ' Show Total
            sheet.ListObjects(0).DisplayTotalRow = True

            ' Column 1 will display "Total" in the totals row
            sheet.ListObjects(0).Columns(0).TotalsRowLabel = "Total"
            ' Column 2 will have no totals calculation
            sheet.ListObjects(0).Columns(1).TotalsCalculation = ExcelTotalsCalculation.None
            ' Column 3 will have no totals calculation
            sheet.ListObjects(0).Columns(2).TotalsCalculation = ExcelTotalsCalculation.None
            ' Column 4 will calculate the sum
            sheet.ListObjects(0).Columns(3).TotalsCalculation = ExcelTotalsCalculation.Sum
            ' Column 5 will calculate the sum
            sheet.ListObjects(0).Columns(4).TotalsCalculation = ExcelTotalsCalculation.Sum
            ' Enable row stripes in the table style
            sheet.ListObjects(0).ShowTableStyleRowStripes = True
            ' Enable column stripes in the table style
            sheet.ListObjects(0).ShowTableStyleColumnStripes = True
            ' Save the modified workbook to a new file named "Sample.xlsx" in Excel 2010 format
            workbook.SaveToFile("Sample.xlsx", ExcelVersion.Version2010)
            ' Release the resources used by the workbook
            workbook.Dispose()
            ExcelDocViewer("Sample.xlsx")
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
