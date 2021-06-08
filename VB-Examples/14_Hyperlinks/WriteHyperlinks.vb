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
			Dim workbook As New Workbook()
			workbook.LoadFromFile("..\..\..\..\..\..\Data\WriteHyperlinks.xlsx")
			Dim sheet As Worksheet = workbook.Worksheets(0)

			sheet.Range("B9").Text = "Home page"
			Dim hylink1 As HyperLink = sheet.HyperLinks.Add(sheet.Range("B10"))
			hylink1.Type = HyperLinkType.Url
			hylink1.Address = "http://www.e-iceblue.com"

			sheet.Range("B11").Text = "Support"
			Dim hylink2 As HyperLink = sheet.HyperLinks.Add(sheet.Range("B12"))
			hylink2.Type = HyperLinkType.Url
			hylink2.Address = "mailto:support@e-iceblue.com"

			sheet.Range("B13").Text = "Forum"
			Dim hylink3 As HyperLink = sheet.HyperLinks.Add(sheet.Range("B14"))
			hylink3.Type = HyperLinkType.Url
			hylink3.Address = "https://www.e-iceblue.com/forum/"

			Dim result As String = "Output_WriteHyperlinks.xlsx"
			workbook.SaveToFile(result, ExcelVersion.Version2010)
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
