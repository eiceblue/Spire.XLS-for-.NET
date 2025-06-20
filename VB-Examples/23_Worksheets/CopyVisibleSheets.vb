Imports Spire.Xls
Imports System.ComponentModel
Imports System.Text

Namespace CopyVisibleSheets
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new Workbook object
            Dim workbook As New Workbook()

            ' Load an existing workbook from a file
            workbook.LoadFromFile("..\..\..\..\..\..\Data\CopyVisibleSheets.xlsx")

            ' Create a new Workbook object for the copied sheets
            Dim workbookNew As New Workbook()
            workbookNew.Version = ExcelVersion.Version2013
            workbookNew.Worksheets.Clear()

            ' Iterate through each worksheet in the original workbook
            For Each sheet As Worksheet In workbook.Worksheets

                ' Check if the worksheet is visible
                If sheet.Visibility = WorksheetVisibility.Visible Then

                    ' Get the name of the visible sheet
                    Dim name As String = sheet.Name

                    ' Add a copy of the visible sheet to the new workbook
                    workbookNew.Worksheets.AddCopy(sheet)
                End If
            Next sheet

            ' Specify the output file name for the new workbook
            Dim result As String = "CopyVisibleSheets_out.xlsx"

            ' Save the new workbook to a file using Excel 2013 format
            workbookNew.SaveToFile(result, ExcelVersion.Version2013)

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
		Private Sub btnClose_Click_1(ByVal sender As Object, ByVal e As EventArgs) Handles btnClose.Click
			Close()
		End Sub
	End Class
End Namespace
