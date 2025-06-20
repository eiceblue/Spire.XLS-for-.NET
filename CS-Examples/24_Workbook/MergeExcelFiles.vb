Imports Spire.Xls
Imports System.ComponentModel
Imports System.Text

Namespace MergeExcelFiles
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
            ' Create a new List to store file paths
            Dim files As New List(Of String)()

            ' Add the first file path to the list
            files.Add("..\..\..\..\..\..\Data\MergeExcelFiles-1.xlsx")

            ' Add the second file path to the list
            files.Add("..\..\..\..\..\..\Data\MergeExcelFiles-2.xls")

            ' Add the third file path to the list
            files.Add("..\..\..\..\..\..\Data\MergeExcelFiles-3.xlsx")

            ' Create a new Workbook object for the merged Excel files
            Dim workbook As New Workbook()

            ' Set the Excel version of the new workbook to Version2013
            workbook.Version = ExcelVersion.Version2013

            ' Clear any existing worksheets in the new workbook
            workbook.Worksheets.Clear()

            ' Create a temporary Workbook object
            Dim tempbook As New Workbook()

            ' Iterate over each file path in the list
            For Each file As String In files
                ' Load the current file into the temporary workbook
                tempbook.LoadFromFile(file)

                ' Iterate over each worksheet in the temporary workbook
                For Each sheet As Worksheet In tempbook.Worksheets
                    ' Copy each worksheet from the temporary workbook to the new workbook
                    workbook.Worksheets.AddCopy(sheet, WorksheetCopyType.CopyAll)
                Next sheet
            Next file

            ' Save the merged workbook to the specified file path with Excel version compatibility set to Version2010
            workbook.SaveToFile("MergeExcelFiles.xlsx", ExcelVersion.Version2010)

            ' Release the resources used by the workbook
            workbook.Dispose()

            ExcelDocViewer("MergeExcelFiles.xlsx")
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
