Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel
Imports Spire.Xls
Imports Spire.Xls.Core.Spreadsheet

Namespace AddPageBreakInXlsFile
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new Workbook object
            Dim workbook As New Workbook()

            ' Load an existing Excel file from a specified path
            workbook.LoadFromFile("..\..\..\..\..\..\Data\Template_Xls_4.xlsx")

            ' Get the first worksheet of the workbook
            Dim sheet As Worksheet = workbook.Worksheets(0)

            ' Add a horizontal page break at cell E4
            sheet.HPageBreaks.Add(sheet.Range("E4"))

            ' Add a vertical page break at cell C4
            sheet.VPageBreaks.Add(sheet.Range("C4"))

            ' Specify the name for the resulting Excel file
            Dim result As String = "Result-AddPageBreakInXlsFile.xlsx"

            ' Save the modified workbook to a new file with Excel 2013 format
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
