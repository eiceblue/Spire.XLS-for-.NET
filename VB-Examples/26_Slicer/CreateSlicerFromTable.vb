Imports Spire.Xls
Imports Spire.Xls.Core
Imports System.ComponentModel
Imports System.Text

Namespace CreateSlicerFromTable
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			Dim wb As New Workbook()
			Dim worksheet As Worksheet = wb.Worksheets(0)
			worksheet.Range("A1").Value = "fruit"
			worksheet.Range("A2").Value = "grape"
			worksheet.Range("A3").Value = "blueberry"
			worksheet.Range("A4").Value = "kiwi"
			worksheet.Range("A5").Value = "cherry"
			worksheet.Range("A6").Value = "grape"
			worksheet.Range("A7").Value = "blueberry"
			worksheet.Range("A8").Value = "kiwi"
			worksheet.Range("A9").Value = "cherry"

			worksheet.Range("B1").Value = "year"
			worksheet.Range("B2").Value2 = 2020
			worksheet.Range("B3").Value2 = 2020
			worksheet.Range("B4").Value2 = 2020
			worksheet.Range("B5").Value2 = 2020
			worksheet.Range("B6").Value2 = 2021
			worksheet.Range("B7").Value2 = 2021
			worksheet.Range("B8").Value2 = 2021
			worksheet.Range("B9").Value2 = 2021

			worksheet.Range("C1").Value = "amount"
			worksheet.Range("C2").Value2 = 50
			worksheet.Range("C3").Value2 = 60
			worksheet.Range("C4").Value2 = 70
			worksheet.Range("C5").Value2 = 80
			worksheet.Range("C6").Value2 = 90
			worksheet.Range("C7").Value2 = 100
			worksheet.Range("C8").Value2 = 110
			worksheet.Range("C9").Value2 = 120

			' Get slicer collection
			Dim slicers As XlsSlicerCollection = worksheet.Slicers

			'Create a table with the data from the specific cell range.
			Dim table As IListObject = worksheet.ListObjects.Create("Super Table", worksheet.Range("A1:C9"))

			Dim count As Integer = 3
			Dim index As Integer = 0
			For Each type As SlicerStyleType In System.Enum.GetValues(GetType(SlicerStyleType))
				count += 5
				Dim range As String = "E" & count
				index = slicers.Add(table, range.ToString(), 0)

				'Style setting
				Dim xlsSlicer As XlsSlicer = slicers(index)
				xlsSlicer.Name = "slicers_" & count
				xlsSlicer.StyleType = type
			Next type

			'Save to file
			wb.SaveToFile("CreateSlicerFromTable.xlsx", ExcelVersion.Version2013)

			' Dispose of the workbook object to release resources
			wb.Dispose()

			' Launch the file
			ExcelDocViewer("CreateSlicerFromTable.xlsx")
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
