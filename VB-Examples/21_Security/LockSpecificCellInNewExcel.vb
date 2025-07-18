Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel
Imports Spire.Xls
Imports Spire.Xls.Core.Spreadsheet

Namespace LockSpecificCellInNewExcel
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new Workbook object
            Dim workbook As New Workbook()

            ' Create an empty sheet in the Workbook
            workbook.CreateEmptySheet()

            ' Get the first worksheet from the Workbook
            Dim sheet As Worksheet = workbook.Worksheets(0)

            ' Loop through rows 0 to 254 and set the Locked property of each row's style to False
            For i As Integer = 0 To 254
                sheet.Rows(i).Style.Locked = False
            Next i

            ' Set the text of cell A1 in the worksheet to "Locked"
            sheet.Range("A1").Text = "Locked"

            ' Set the Locked property of cell A1's style to True
            sheet.Range("A1").Style.Locked = True

            ' Set the text of cells C1 to E3 in the worksheet to "Locked"
            sheet.Range("C1:E3").Text = "Locked"

            ' Set the Locked property of cells C1 to E3's style to True
            sheet.Range("C1:E3").Style.Locked = True

            ' Protect the worksheet with the specified password and enable all protection options
            sheet.Protect("123", SheetProtectionType.All)

            ' Specify the name of the resulting Excel file after locking specific cells
            Dim result As String = "Result-LockSpecificCellInNewlyXlsFile.xlsx"

            ' Save the Workbook to the specified path in Excel 2013 format
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
