Imports Spire.Xls
Imports System.ComponentModel
Imports System.Text

Namespace ProtectWithEditableRange
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new Workbook object
            Dim workbook As New Workbook()

            ' Load the Excel file from the specified path
            workbook.LoadFromFile("..\..\..\..\..\..\Data\ProtectWithEditableRange.xlsx")

            ' Get the first worksheet from the Workbook
            Dim sheet As Worksheet = workbook.Worksheets(0)

            ' Add an editable range named "EditableRanges" for the specified range B4:E12 in the worksheet
            sheet.AddAllowEditRange("EditableRanges", sheet.Range("B4:E12"))

            ' Protect the worksheet with the specified password and enable all protection options
            sheet.Protect("TestPassword", SheetProtectionType.All)

            ' Specify the name of the resulting Excel file after protecting with editable range
            Dim result As String = "ProtectWithEditableRange_result.xlsx"

            ' Save the Workbook to the specified path in Excel 2010 format
            workbook.SaveToFile(result, ExcelVersion.Version2010)

            ' Release the resources used by the workbook
            workbook.Dispose()

            ExcelDocViewer(result)
		End Sub

		Private Sub ExcelDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub
		Private Sub btnAbout_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnAbout.Click
			Close()
		End Sub
	End Class
End Namespace
