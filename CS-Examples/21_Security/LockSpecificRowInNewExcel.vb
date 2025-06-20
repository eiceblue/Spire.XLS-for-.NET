Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel
Imports Spire.Xls
Imports Spire.Xls.Core.Spreadsheet

Namespace LockSpecificRowInNewExcel
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new instance of Workbook
            Dim workbook As New Workbook()

            ' Create a new empty sheet in the workbook
            workbook.CreateEmptySheet()

            ' Get the first worksheet from the workbook
            Dim sheet As Worksheet = workbook.Worksheets(0)

            ' Loop through rows 0 to 254 and set the Locked property to False for each row
            For i As Integer = 0 To 254
                sheet.Rows(i).Style.Locked = False
            Next i

            ' Set the text of Row 2 to "Locked" and set its Locked property to True
            sheet.Rows(2).Text = "Locked"
            sheet.Rows(2).Style.Locked = True

            ' Protect the worksheet with a password and specify SheetProtectionType.All (all elements are protected)
            sheet.Protect("123", SheetProtectionType.All)

            ' Specify the name of the resulting file
            Dim result As String = "Result-LockSpecificRowInNewlyXlsFile.xlsx"

            ' Save the modified workbook to a new file with the specified name and Excel version (2013)
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
