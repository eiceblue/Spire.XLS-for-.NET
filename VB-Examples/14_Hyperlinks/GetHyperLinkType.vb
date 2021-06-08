Imports System.IO
Imports System.Text
Imports Spire.Xls

Namespace GetHyperLinkType

	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a workbook
			Dim workbook As New Workbook()

			'Load the document from disk
			workbook.LoadFromFile("..\..\..\..\..\..\Data\HyperlinksSample2.xlsx")

			'Get the first worksheet
			Dim sheet As Worksheet = workbook.Worksheets(0)

			'Iterate all hyperlinks
			Dim sb As New StringBuilder()
			For Each item In sheet.HyperLinks
				'Get hyperlink address
				Dim address As String = item.Address
				'Get hyperlink type
				Dim type As HyperLinkType = item.Type
				sb.AppendLine("Link address: " & address)
				sb.AppendLine("Link type: " & type.ToString())
				sb.AppendLine()
			Next item

			'Save to Text file
			Dim output As String = "GetHyperLinkType.txt"
			File.WriteAllText(output, sb.ToString())

			'Launch the file
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
