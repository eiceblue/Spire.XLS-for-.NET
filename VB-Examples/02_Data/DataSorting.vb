Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel

Imports Spire.Xls

Namespace DataSorting
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new Workbook object
            Dim workbook As New Workbook()

            ' Load an existing Excel file into the workbook
            workbook.LoadFromFile("..\..\..\..\..\..\Data\DataSorting.xls")

            ' Get the first worksheet from the workbook
            Dim worksheet As Worksheet = workbook.Worksheets(0)

            ' Add sort columns for sorting by column 2 (ascending) and column 3 (ascending)
            workbook.DataSorter.SortColumns.Add(2, OrderBy.Ascending)
            workbook.DataSorter.SortColumns.Add(3, OrderBy.Ascending)

            ' Perform the sorting operation on the specified range of cells in the worksheet (A1 to E19)
            workbook.DataSorter.Sort(worksheet.Range("A1:E19"))

            ' Specify the name for the resulting Excel file
            Dim result As String = "DataSorting_out.xlsx"

            ' Save the modified workbook to the specified file path in Excel 2013 format
            workbook.SaveToFile(result, ExcelVersion.Version2013)
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

		Private Sub Form1_Load(ByVal sender As Object, ByVal e As EventArgs) Handles MyBase.Load

		End Sub

	End Class
End Namespace
