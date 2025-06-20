Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel
Imports Spire.Xls
Imports Spire.Xls.Core.Spreadsheet
Imports System.Text
Imports System.IO

Namespace GetExcelPaperDimensions
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new Workbook object
            Dim workbook As New Workbook()

            ' Get the first worksheet from the workbook
            Dim sheet As Worksheet = workbook.Worksheets(0)

            ' Create a StringBuilder object to store the content
            Dim content As New StringBuilder()

            ' Set the paper size of the worksheet to A2 and append the dimensions to the content
            sheet.PageSetup.PaperSize = PaperSizeType.A2Paper
            content.AppendLine("A2Paper: " & sheet.PageSetup.PageWidth & " x " & sheet.PageSetup.PageHeight)

            ' Set the paper size of the worksheet to A3 and append the dimensions to the content
            sheet.PageSetup.PaperSize = PaperSizeType.PaperA3
            content.AppendLine("PaperA3: " & sheet.PageSetup.PageWidth & " x " & sheet.PageSetup.PageHeight)

            ' Set the paper size of the worksheet to A4 and append the dimensions to the content
            sheet.PageSetup.PaperSize = PaperSizeType.PaperA4
            content.AppendLine("PaperA4: " & sheet.PageSetup.PageWidth & " x " & sheet.PageSetup.PageHeight)

            ' Set the paper size of the worksheet to Letter and append the dimensions to the content
            sheet.PageSetup.PaperSize = PaperSizeType.PaperLetter
            content.AppendLine("PaperLetter: " & sheet.PageSetup.PageWidth & " x " & sheet.PageSetup.PageHeight)

            ' Specify the filename for the resulting text file
            Dim result As String = "Result-GetExcelPaperDimensions.txt"

            ' Write the content to the specified filename
            File.WriteAllText(result, content.ToString())

            ' Release the resources used by the workbook
            workbook.Dispose()

            'Launch the file.
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
