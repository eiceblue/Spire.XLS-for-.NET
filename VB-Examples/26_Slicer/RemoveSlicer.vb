Imports Spire.Xls
Imports Spire.Xls.Core
Imports System.ComponentModel
Imports System.Text

Namespace RemoveSlicer
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			' Create a new Workbook instance
			Dim wb As New Workbook()

			' Load an existing Excel file from the specified path
			wb.LoadFromFile("..\..\..\..\..\..\Data\SlicerTemplate.xlsx")

			' Get the first worksheet in the workbook
			Dim worksheet As Worksheet = wb.Worksheets(0)

			' Get the slicer collection from the worksheet
			Dim slicers As XlsSlicerCollection = worksheet.Slicers

			' Example: Remove the first slicer in the collection 
			' slicers.RemoveAt(0);

			' Clear all slicers from the collection
			slicers.Clear()

			' Save the modified workbook to a new file with Excel 2013 version format
			wb.SaveToFile("RemoveSlicer.xlsx", ExcelVersion.Version2013)

			' Dispose of the workbook object to release resources
			wb.Dispose()

			' Launch the file
			ExcelDocViewer("RemoveSlicer.xlsx")
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
