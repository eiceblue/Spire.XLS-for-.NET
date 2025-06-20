Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel

Imports Spire.Xls
Imports Spire.Xls.Charts

Namespace ProtectCell
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new Workbook object
            Dim workbook As New Workbook()
            ' Load the workbook from a specific file path
            workbook.LoadFromFile("..\..\..\..\..\..\Data\ProtectCell.xlsx")

            ' Get the first worksheet in the workbook
            Dim sheet As Worksheet = workbook.Worksheets(0)
            ' Set the Locked property of the entire allocated range to false, allowing changes
            sheet.AllocatedRange.Style.Locked = False

            ' Set the Locked property of cell B3 to true, preventing changes
            sheet.Range("B3").Style.Locked = True

            ' Protect the worksheet with a password using SheetProtectionType.All option
            sheet.Protect("TestPassword", SheetProtectionType.All)

            ' Define the output file name for the protected workbook
            Dim result As String = "ProtectCell_result.xlsx"
            ' Save the protected workbook to the specified file path and Excel version (2013)
            workbook.SaveToFile(result, ExcelVersion.Version2013)

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

		Private Sub btnClose_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnClose.Click
			Close()
		End Sub
	End Class
End Namespace
