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
            ' Create a new Workbook object
            Dim workbook As New Workbook()

            ' Load the Excel file from a specified path
            workbook.LoadFromFile("..\..\..\..\..\..\Data\HyperlinksSample2.xlsx")

            ' Get the first worksheet from the workbook
            Dim sheet As Worksheet = workbook.Worksheets(0)

            ' Create a StringBuilder object to store the hyperlinks information
            Dim sb As New StringBuilder()

            ' Iterate through each hyperlink in the worksheet
            For Each item In sheet.HyperLinks

                ' Get the address of the hyperlink
                Dim address As String = item.Address

                ' Get the type of the hyperlink
                Dim type As HyperLinkType = item.Type

                ' Append the link address and type to the StringBuilder object
                sb.AppendLine("Link address: " & address)
                sb.AppendLine("Link type: " & type.ToString())
                sb.AppendLine()

            Next item

            ' Specify the output file name
            Dim output As String = "GetHyperLinkType.txt"

            ' Write the content of the StringBuilder object to a text file
            File.WriteAllText(output, sb.ToString())

            ' Release the resources used by the workbook
            workbook.Dispose()

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
