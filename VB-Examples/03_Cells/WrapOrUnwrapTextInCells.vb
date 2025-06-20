Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel
Imports Spire.Xls
Imports Spire.Xls.Core.Spreadsheet

Namespace WrapOrUnwrapTextInCells
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            'Creates a new instance of the Workbook class.
            Dim workbook As New Workbook()

            'Retrieves the first worksheet from the workbook.
            Dim sheet As Worksheet = workbook.Worksheets(0)

            'Sets the text value for cell "C1".
            sheet.Range("C1").Text = "e-iceblue is in facebook and welcome to like us"
            'Enables text wrapping for cell "C1".
            sheet.Range("C1").Style.WrapText = True
            'Sets the text value for cell "D1".
            sheet.Range("D1").Text = "e-iceblue is in twitter and welcome to follow us"
            'Enables text wrapping for cell "D1".
            sheet.Range("D1").Style.WrapText = True

            'Sets the text value for cell "C2".
            sheet.Range("C2").Text = "http://www.facebook.com/pages/e-iceblue/139657096082266"
            'Disables text wrapping for cell "C2".
            sheet.Range("C2").Style.WrapText = False
            'Sets the text value for cell "D2".
            sheet.Range("D2").Text = "https://twitter.com/eiceblue"
            'Disables text wrapping for cell "D2".
            sheet.Range("D2").Style.WrapText = False

            'Sets the font size for range "C1:D1".
            sheet.Range("C1:D1").Style.Font.Size = 15
            'Sets the font color for range "C1:D1" to blue.
            sheet.Range("C1:D1").Style.Font.Color = Color.Blue

            'Sets the font size for range "C2:D2".
            sheet.Range("C2:D2").Style.Font.Size = 15
            'Sets the font color for range "C2:D2" to deep sky blue.
            sheet.Range("C2:D2").Style.Font.Color = Color.DeepSkyBlue
            'Specifies the name of the output file.
            Dim result As String = "Result-WrapOrUnwrapTextInExcelCells.xlsx"

            'Saves the workbook to the specified file in Excel 2013 format.
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
