Imports Spire.Xls
Imports System.ComponentModel
Imports System.Text

Namespace SetHeaderFooter
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Load a Workbook from disk        
			Dim Workbook As New Workbook()
			Workbook.LoadFromFile("..\..\..\..\..\..\Data\SetHeaderFooter.xlsx")

			'Get the first worksheet
			Dim Worksheet As Worksheet = Workbook.Worksheets(0)


			'Set left header,"Arial Unicode MS" is font name, "18" is font size.
			Worksheet.PageSetup.LeftHeader = "&""Arial Unicode MS""&14 Spire.XLS for .NET "

			'Set center footer 
			Worksheet.PageSetup.CenterFooter = "Footer Text"

			Worksheet.ViewMode = ViewMode.Layout

			Dim result As String = "SetHeaderFooter_result.xlsx"
			'Save and Launch
			Workbook.SaveToFile(result, ExcelVersion.Version2010)
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
