Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel

Imports Spire.Xls

Namespace WriteHyperlinks

	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new workbook object
            Dim workbook As New Workbook()

            ' Load an existing workbook from the specified file path
            workbook.LoadFromFile("..\..\..\..\..\..\Data\WriteHyperlinks.xlsx")

            ' Get the first worksheet in the workbook
            Dim sheet As Worksheet = workbook.Worksheets(0)

            ' Set the text of cell B9 as "Home page"
            sheet.Range("B9").Text = "Home page"

            ' Create a hyperlink object and associate it with cell B10
            Dim hylink1 As HyperLink = sheet.HyperLinks.Add(sheet.Range("B10"))

            ' Set the type of the hyperlink to a URL
            hylink1.Type = HyperLinkType.Url

            ' Set the address of the hyperlink to "http://www.e-iceblue.com"
            hylink1.Address = "http://www.e-iceblue.com"

            ' Set the text of cell B11 as "Support"
            sheet.Range("B11").Text = "Support"

            ' Create a hyperlink object and associate it with cell B12
            Dim hylink2 As HyperLink = sheet.HyperLinks.Add(sheet.Range("B12"))

            ' Set the type of the hyperlink to a URL
            hylink2.Type = HyperLinkType.Url

            ' Set the address of the hyperlink to "mailto:support@e-iceblue.com"
            hylink2.Address = "mailto:support@e-iceblue.com"

            ' Set the text of cell B13 as "Forum"
            sheet.Range("B13").Text = "Forum"

            ' Create a hyperlink object and associate it with cell B14
            Dim hylink3 As HyperLink = sheet.HyperLinks.Add(sheet.Range("B14"))

            ' Set the type of the hyperlink to a URL
            hylink3.Type = HyperLinkType.Url

            ' Set the address of the hyperlink to "https://www.e-iceblue.com/forum/"
            hylink3.Address = "https://www.e-iceblue.com/forum/"

            ' Define the output file name as "Output_WriteHyperlinks.xlsx"
            Dim result As String = "Output_WriteHyperlinks.xlsx"

            ' Save the modified workbook to the specified file path using Excel 2010 format
            workbook.SaveToFile(result, ExcelVersion.Version2010)

            ' Release the resources used by the workbook
            workbook.Dispose()

            ExcelDocViewer(result)
		End Sub

		Private Sub btnClose_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnClose.Click
			Close()
		End Sub
		Private Sub ExcelDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub

	End Class
End Namespace
