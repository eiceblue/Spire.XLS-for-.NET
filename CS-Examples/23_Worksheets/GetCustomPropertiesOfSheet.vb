Imports Spire.Xls
Imports Spire.Xls.Core
Imports System.IO
Imports Spire.Xls.Core.Spreadsheet.Collections

Namespace GetCustomPropertiesOfSheet
	Partial Public Class Form1
		Inherits Form

		Public Sub New()
			InitializeComponent()
		End Sub

		 Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			' Create a new workbook
			Dim workbook As New Workbook()

			' Load a Workbook from disk
			workbook.LoadFromFile("..\..\..\..\..\..\Data\GetCustomPropertiesOfSheet.xlsx")

			' Get the first sheet
			Dim worksheet As Worksheet = workbook.Worksheets(0)

			' Get the custom properties of the first sheet
			Dim customProperties As ICustomPropertiesCollection = worksheet.CustomProperties
			Dim information As String = ""
			For i As Integer = 0 To customProperties.Count - 1
				Dim xcp As XlsCustomProperty = customProperties(i)
				Dim name As String = xcp.Name
				information &= "Name:" & name & vbLf
				Dim value As String = xcp.Value
				information &= "Value:" & value & vbLf
			Next i

			' Save the information to a .txt file
			Dim result As String = "GetCustomPropertiesOfSheet-out.txt"
			File.WriteAllText(result, information)

			' Dispose of the workbook object to release resources
			workbook.Dispose()

			' Launch the file
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
