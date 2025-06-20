Imports Spire.Xls
Imports System.Text
Imports System.IO

Namespace GetPageCount

	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new workbook object
            Dim workbook As New Workbook()

            ' Load the Excel file from the specified path
            workbook.LoadFromFile("..\..\..\..\..\..\Data\WorksheetSample2.xlsx")

            ' Retrieve the split page information for each worksheet in the workbook
            Dim pageInfoList = workbook.GetSplitPageInfo()

            ' Create a StringBuilder object to store the output text
            Dim stringBuilder As New StringBuilder()

            ' Iterate through each worksheet in the workbook
            For i As Integer = 0 To workbook.Worksheets.Count - 1
                ' Retrieve the name of the current worksheet
                Dim sheetname As String = workbook.Worksheets(i).Name

                ' Retrieve the page count for the current worksheet
                Dim pagecount As Integer = pageInfoList(i).Count

                ' Append the worksheet name and its page count to the StringBuilder
                stringBuilder.AppendLine(sheetname & "'s page count is: " & pagecount)
            Next i

            ' Specify the output file path and name
            Dim output As String = "GetPageCount.txt"

            ' Write the contents of the StringBuilder to the output file
            File.WriteAllText(output, stringBuilder.ToString())

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
