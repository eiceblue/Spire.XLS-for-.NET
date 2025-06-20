Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel
Imports System.IO
Imports Spire.Xls
Imports Spire.Xls.Charts
Imports System.Text

Namespace FillDataInWorksheet
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new workbook
            Dim workbook As New Workbook()

            ' Get the first worksheet in the workbook
            Dim worksheet As Worksheet = workbook.Worksheets(0)

            ' Set the font of cells A1, B1, and C1 to bold
            worksheet.Range("A1").Style.Font.IsBold = True
            worksheet.Range("B1").Style.Font.IsBold = True
            worksheet.Range("C1").Style.Font.IsBold = True

            ' Set the text of cell A1 to "Month"
            worksheet.Range("A1").Text = "Month"

            ' Set the texts of cells A2 to A5 with specific months
            worksheet.Range("A2").Text = "January"
            worksheet.Range("A3").Text = "February"
            worksheet.Range("A4").Text = "March"
            worksheet.Range("A5").Text = "April"

            ' Set the text of cell B1 to "Payments"
            worksheet.Range("B1").Text = "Payments"

            ' Set the numeric values of cells B2 to B5 with specific payment values
            worksheet.Range("B2").NumberValue = 251
            worksheet.Range("B3").NumberValue = 515
            worksheet.Range("B4").NumberValue = 454
            worksheet.Range("B5").NumberValue = 874

            ' Set the text of cell C1 to "Sample"
            worksheet.Range("C1").Text = "Sample"

            ' Set the texts of cells C2 to C5 with specific sample values
            worksheet.Range("C2").Text = "Sample1"
            worksheet.Range("C3").Text = "Sample2"
            worksheet.Range("C4").Text = "Sample3"
            worksheet.Range("C5").Text = "Sample4"

            ' Set the width of column 2 (B) to 10
            worksheet.SetColumnWidth(2, 10)

            ' Specify the output file name as "Output.xlsx"
            Dim outputFile As String = "Output.xlsx"

            ' Save the workbook to the specified file in Excel 2013 format
            workbook.SaveToFile(outputFile, ExcelVersion.Version2013)

            ' Release the resources used by the workbook
            workbook.Dispose()

            'Launching the output file.
            Viewer(outputFile)
		End Sub
		Private Sub Viewer(ByVal fileName As String)
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
