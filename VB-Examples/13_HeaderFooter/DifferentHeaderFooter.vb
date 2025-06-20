Imports Spire.Xls
Imports System.ComponentModel
Imports System.Text

Namespace DifferentHeaderFooter
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new Workbook object
            Dim workbook As New Workbook()

            ' Load a workbook from a specified file path
            workbook.LoadFromFile("..\..\..\..\..\..\Data\DifferentHeaderFooter.xlsx")

            ' Get the first worksheet from the workbook
            Dim sheet As Worksheet = workbook.Worksheets(0)

            ' Set the text in cell A1 to "Page 1"
            sheet.Range("A1").Text = "Page 1"

            ' Set the text in cell G1 to "Page 2"
            sheet.Range("G1").Text = "Page 2"

            ' Enable different odd and even page headers and footers
            sheet.PageSetup.DifferentOddEven = 1

            ' Set the odd page header text format
            sheet.PageSetup.OddHeaderString = "&""Arial""&12&B&KFFC000 Odd_Header"

            ' Set the odd page footer text format
            sheet.PageSetup.OddFooterString = "&""Arial""&12&B&KFFC000 Odd_Footer"

            ' Set the even page header text format
            sheet.PageSetup.EvenHeaderString = "&""Arial""&12&B&KFF0000 Even_Header"

            ' Set the even page footer text format
            sheet.PageSetup.EvenFooterString = "&""Arial""&12&B&KFF0000 Even_Footer"

            ' Change the view mode of the worksheet to Layout view
            sheet.ViewMode = ViewMode.Layout

            ' Save the workbook to a file named "Output.xlsx" in Excel 2013 format
            workbook.SaveToFile("Output.xlsx", ExcelVersion.Version2013)

            ' Release the resources used by the workbook
            workbook.Dispose()
            ExcelDocViewer("Output.xlsx")
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
