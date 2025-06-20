Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel

Imports Spire.Xls

Namespace TextAlign

	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new workbook object
            Dim workbook As New Workbook()

            ' Load the workbook from a file
            workbook.LoadFromFile("..\..\..\..\..\..\Data\TextAlign.xlsx")

            ' Get the first worksheet in the workbook
            Dim sheet As Worksheet = workbook.Worksheets(0)

            ' Set vertical alignment of range B1:C1 to Top
            sheet.Range("B1:C1").Style.VerticalAlignment = VerticalAlignType.Top

            ' Set vertical alignment of range B2:C2 to Center
            sheet.Range("B2:C2").Style.VerticalAlignment = VerticalAlignType.Center

            ' Set vertical alignment of range B3:C3 to Bottom
            sheet.Range("B3:C3").Style.VerticalAlignment = VerticalAlignType.Bottom

            ' Set horizontal alignment of range B4:C4 to General
            sheet.Range("B4:C4").Style.HorizontalAlignment = HorizontalAlignType.General

            ' Set horizontal alignment of range B5:C5 to Left
            sheet.Range("B5:C5").Style.HorizontalAlignment = HorizontalAlignType.Left

            ' Set horizontal alignment of range B6:C6 to Center
            sheet.Range("B6:C6").Style.HorizontalAlignment = HorizontalAlignType.Center

            ' Set horizontal alignment of range B7:C7 to Right
            sheet.Range("B7:C7").Style.HorizontalAlignment = HorizontalAlignType.Right

            ' Set rotation angle of range B8:C8 to 45 degrees
            sheet.Range("B8:C8").Style.Rotation = 45

            ' Set rotation angle of range B9:C9 to 90 degrees
            sheet.Range("B9:C9").Style.Rotation = 90

            ' Set row height of range B8:C9 to 60
            sheet.Range("B8:C9").RowHeight = 60

            ' Save the modified workbook to a new file
            Dim result As String = "Result-TextAlign.xlsx"
            workbook.SaveToFile(result, ExcelVersion.Version2010)

            ' Release the resources used by the workbook
            workbook.Dispose()

            'Launch the Excel file
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
