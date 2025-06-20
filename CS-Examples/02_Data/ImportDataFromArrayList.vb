Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel

Imports Spire.Xls

Namespace ImportDataFromArrayList
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new workbook object
            Dim workbook As New Workbook()

            ' Create an empty worksheet within the workbook
            workbook.CreateEmptySheets(1)

            ' Get the first worksheet from the workbook
            Dim sheet As Worksheet = workbook.Worksheets(0)

            ' Create an ArrayList object
            Dim list As New ArrayList()

            ' Add strings to the ArrayList
            list.Add("Spire.Doc for .NET")
            list.Add("Spire.XLS for .NET")
            list.Add("Spire.PDF for .NET")
            list.Add("Spire.Presentation for .NET")

            ' Insert the ArrayList data into the worksheet starting at cell position (1, 1)
            sheet.InsertArrayList(list, 1, 1, True)

            ' Save the workbook to an Excel file named "ImportDataFromArrayList_out.xlsx" in Excel 2013 format
            Dim result As String = "ImportDataFromArrayList_out.xlsx"
            workbook.SaveToFile(result, ExcelVersion.Version2013)
            ' Release the resources used by the workbook
            workbook.Dispose()

            'Launch the Excel file
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
