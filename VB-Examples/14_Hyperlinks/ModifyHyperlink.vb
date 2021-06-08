Imports Spire.Xls
Imports Spire.Xls.Collections

Namespace ModifyHyperlink

	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a workbook
			Dim workbook As New Workbook()

			'Load the document from disk
			workbook.LoadFromFile("..\..\..\..\..\..\Data\ModifyHyperlink.xlsx")

			'Get the collection of all hyperlinks in the worksheet
			Dim sheet As Worksheet = workbook.Worksheets(0)

			'Change the values of TextToDisplay and Address property 
			Dim links As HyperLinksCollection = sheet.HyperLinks
			links(0).TextToDisplay = "Spire.XLS for .NET"
			links(0).Address = "http://www.e-iceblue.com/Introduce/excel-for-net-introduce.html"

			'Save the document
			Dim output As String = "ModifyHyperlinkResult.xlsx"
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
