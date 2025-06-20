Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel
Imports Spire.Xls
Imports Spire.Xls.Core.Spreadsheet
Imports Spire.Xls.Core

Namespace AddTotalRowToTable
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new Excel workbook object.
            Dim workbook As New Workbook()

            ' Load an existing Excel file from the specified path.
            workbook.LoadFromFile("..\..\..\..\..\..\Data\AddATotalRowToTable.xlsx")

            ' Get the first worksheet in the workbook.
            Dim sheet As Worksheet = workbook.Worksheets(0)

            ' Create a table named "Table" using the data from cells A1 to D4.
            Dim table As IListObject = sheet.ListObjects.Create("Table", sheet.Range("A1:D4"))

            ' Enable the display of the total row for the table.
            table.DisplayTotalRow = True

            ' Set the label for the total row in column 0 as "Total".
            table.Columns(0).TotalsRowLabel = "Total"
            ' Set the calculation method for totals in column 1 as sum.
            table.Columns(1).TotalsCalculation = ExcelTotalsCalculation.Sum
            ' Set the calculation method for totals in column 2 as sum.
            table.Columns(2).TotalsCalculation = ExcelTotalsCalculation.Sum
            ' Set the calculation method for totals in column 3 as sum.
            table.Columns(3).TotalsCalculation = ExcelTotalsCalculation.Sum
            ' Specify the output file name.
            Dim result As String = "Result-AddATotalRowToTable.xlsx"

            ' Save the modified workbook to the specified file with Excel 2013 format.
            workbook.SaveToFile(result, ExcelVersion.Version2013)
            ' Release the resources used by the workbook
            workbook.Dispose()

            'Launch the MS Excel file.
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
