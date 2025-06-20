Imports Spire.Xls

Namespace AddHyperlinkToText

	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new Workbook object
            Dim workbook As New Workbook()

            ' Load an existing Excel file from the specified path
            workbook.LoadFromFile("..\..\..\..\..\..\Data\CommonTemplate1.xlsx")

            ' Get the first worksheet from the loaded workbook
            Dim sheet As Worksheet = workbook.Worksheets(0)

            ' Add a hyperlink to a cell (D10) in the worksheet
            Dim UrlLink As HyperLink = sheet.HyperLinks.Add(sheet.Range("D10"))
            UrlLink.TextToDisplay = sheet.Range("D10").Text
            UrlLink.Type = HyperLinkType.Url
            UrlLink.Address = "http://en.wikipedia.org/wiki/Chicago"

            ' Add another hyperlink to a cell (E10) in the worksheet
            Dim MailLink As HyperLink = sheet.HyperLinks.Add(sheet.Range("E10"))
            MailLink.TextToDisplay = sheet.Range("E10").Text
            MailLink.Type = HyperLinkType.Url
            MailLink.Address = "mailto:Amor.Aqua@gmail.com"

            ' Specify the filename for the resulting Excel file
            Dim output As String = "AddHyperlinkToText.xlsx"

            ' Save the Workbook object to the specified file path in Excel 2013 format
            workbook.SaveToFile(output, ExcelVersion.Version2013)

            ' Release the resources used by the workbook
            workbook.Dispose()

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
