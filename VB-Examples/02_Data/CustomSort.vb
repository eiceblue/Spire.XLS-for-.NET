Imports Spire.Xls
Imports System.ComponentModel
Imports System.Text

Namespace CustomSort
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new Workbook object
            Dim wb As New Workbook()

            ' Get the first sheet from the workbook
            Dim sheet As Worksheet = wb.Worksheets(0)

            ' Specify that the header should not be included in the sorting process
            wb.DataSorter.IsIncludeTitle = False

            ' Add data to specific cells in the worksheet
            sheet.Range("A1").Text = "AA"
            sheet.Range("A2").Text = "BB"
            sheet.Range("A3").Text = "CC"
            sheet.Range("A4").Text = "DD"
            sheet.Range("A5").Text = "EE"
            sheet.Range("A6").Text = "FF"
            sheet.Range("A7").Text = "GG"
            sheet.Range("A8").Text = "HH"

            ' Configure custom sorting by adding sort columns and specifying the desired order
            wb.DataSorter.SortColumns.Add(0, New String() {"DD", "CC", "BB", "AA", "HH", "GG", "FF", "EE"})

            ' Perform the sorting operation on the specified range of cells in the worksheet
            wb.DataSorter.Sort(wb.Worksheets(0).Range("A1:A8"))

            ' Specify the name for the resulting Excel file
            Dim result As String = "result.xlsx"

            ' Save the modified workbook to the specified file path in Excel 2013 format
            wb.SaveToFile(result, ExcelVersion.Version2013)

            ' Release the resources used by the workbook
            wb.Dispose()

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
