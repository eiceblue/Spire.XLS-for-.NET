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
            ' Create a new Workbook object
            Dim Workbook As New Workbook()

            ' Load the Excel file from the specified path
            Workbook.LoadFromFile("..\..\..\..\..\..\Data\SetHeaderFooter.xlsx")

            ' Get the first worksheet from the Workbook
            Dim Worksheet As Worksheet = Workbook.Worksheets(0)

            ' Set the left header of the page to a specific text
            Worksheet.PageSetup.LeftHeader = "&""Arial Unicode MS""&14 Spire.XLS for .NET"

            ' Set the center footer of the page to a specific text
            Worksheet.PageSetup.CenterFooter = "Footer Text"

            ' Set the view mode of the worksheet to layout
            Worksheet.ViewMode = ViewMode.Layout

            ' Specify the filename for the resulting Excel file
            Dim result As String = "SetHeaderFooter_result.xlsx"

            ' Save the Workbook object to the specified file path in Excel 2010 format
            Workbook.SaveToFile(result, ExcelVersion.Version2010)

            ' Release the resources used by the workbook
            Workbook.Dispose()
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
