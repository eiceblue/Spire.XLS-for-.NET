Imports Spire.Xls

Namespace AddHyperlinkToText

	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a workbook
			Dim workbook As New Workbook()

			'Load the document from disk
			workbook.LoadFromFile("..\..\..\..\..\..\Data\CommonTemplate1.xlsx")

			'Get the first sheet
			Dim sheet As Worksheet = workbook.Worksheets(0)

			'Add url link
			Dim UrlLink As HyperLink = sheet.HyperLinks.Add(sheet.Range("D10"))
			UrlLink.TextToDisplay = sheet.Range("D10").Text
			UrlLink.Type = HyperLinkType.Url
			UrlLink.Address = "http://en.wikipedia.org/wiki/Chicago"

			'Add email link
			Dim MailLink As HyperLink = sheet.HyperLinks.Add(sheet.Range("E10"))
			MailLink.TextToDisplay = sheet.Range("E10").Text
			MailLink.Type = HyperLinkType.Url
			MailLink.Address = "mailto:Amor.Aqua@gmail.com"

			'Save the document
			Dim output As String = "AddHyperlinkToText.xlsx"
			workbook.SaveToFile(output, ExcelVersion.Version2013)

			'Launch the Excel file
			ExcelDocViewer(output)
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
