Imports Spire.Xls
Imports Spire.Xls.Collections

Namespace ModifyHyperlink

	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new workbook object
            Dim workbook As New Workbook()

            ' Load an existing workbook from the specified file path
            workbook.LoadFromFile("..\..\..\..\..\..\Data\ModifyHyperlink.xlsx")

            ' Get the first worksheet in the workbook
            Dim sheet As Worksheet = workbook.Worksheets(0)

            ' Get the collection of hyperlinks in the worksheet
            Dim links As HyperLinksCollection = sheet.HyperLinks

            ' Update the display text and address of the first hyperlink in the collection
            links(0).TextToDisplay = "Spire.XLS for .NET"
            links(0).Address = "http://www.e-iceblue.com/Introduce/excel-for-net-introduce.html"

            ' Define the output file name as "ModifyHyperlinkResult.xlsx"
            Dim output As String = "ModifyHyperlinkResult.xlsx"

            ' Save the modified workbook to the specified file path using Excel 2013 format
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
