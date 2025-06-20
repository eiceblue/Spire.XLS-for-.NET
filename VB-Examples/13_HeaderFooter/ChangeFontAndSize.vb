Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel
Imports Spire.Xls
Imports Spire.Xls.Core.Spreadsheet

Namespace ChangeFontAndSize
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new workbook object
            Dim workbook As New Workbook()

            ' Load an existing workbook from a file
            workbook.LoadFromFile("..\..\..\..\..\..\Data\ChangeFontAndSizeForHeaderAndFooter.xlsx")

            ' Get the first worksheet in the workbook
            Dim sheet As Worksheet = workbook.Worksheets(0)

            ' Get the current text in the left header
            Dim text As String = sheet.PageSetup.LeftHeader

            ' Update the left header text with a custom string and font size
            text = "&""Arial Unicode MS""&18 Header Footer Sample by Spire.XLS "
            sheet.PageSetup.LeftHeader = text

            ' Set the right footer to the same text as the left header
            sheet.PageSetup.RightFooter = text

            ' Specify the output filename for the modified workbook
            Dim result As String = "Result-ChangeFontAndSizeForHeaderAndFooter.xlsx"

            ' Save the modified workbook to a file
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
