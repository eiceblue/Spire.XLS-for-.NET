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
			'Create a Workbook
			Dim wb As New Workbook()
			wb.LoadFromFile("..\..\..\..\..\..\Data\DifferentHeaderFooter.xlsx")

			Dim sheet As Worksheet = wb.Worksheets(0)

			'set text in range
			sheet.Range("A1").Text = "Page 1"
			sheet.Range("G1").Text = "Page 2"

			'Set the different header footer for Odd and Even pages
			sheet.PageSetup.DifferentOddEven = 1

			'Set the header with font, size, bold and color
			sheet.PageSetup.OddHeaderString = "&""Arial""&12&B&KFFC000 Odd_Header"
			sheet.PageSetup.OddFooterString = "&""Arial""&12&B&KFFC000 Odd_Footer"
			sheet.PageSetup.EvenHeaderString = "&""Arial""&12&B&KFF0000 Even_Header"
			sheet.PageSetup.EvenFooterString = "&""Arial""&12&B&KFF0000 Even_Footer"

			sheet.ViewMode = ViewMode.Layout

			'Save and Launch
			wb.SaveToFile("Output.xlsx", ExcelVersion.Version2013)
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
