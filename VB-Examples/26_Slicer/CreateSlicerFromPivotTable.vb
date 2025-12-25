Imports Spire.Xls
Imports Spire.Xls.Core
Imports System.ComponentModel
Imports System.Text

Namespace CreateSlicerFromPivotTable
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

			' Get pivot table collection
			Dim pivotTables As Spire.Xls.Collections.PivotTablesCollection = worksheet.PivotTables

			'Add a PivotTable to the worksheet
			Dim dataRange As CellRange = worksheet.Range("A1:C9")
			Dim cache As PivotCache = wb.PivotCaches.Add(dataRange)

			'Cell to put the pivot table
			Dim pt As Spire.Xls.PivotTable = worksheet.PivotTables.Add("TestPivotTable", worksheet.Range("A12"), cache)

			'Drag the fields to the row area.
			Dim pf As PivotField = TryCast(pt.PivotFields("fruit"), PivotField)
			pf.Axis = AxisTypes.Row
			Dim pf2 As PivotField = TryCast(pt.PivotFields("year"), PivotField)
			pf2.Axis = AxisTypes.Column

			'Drag the field to the data area.
			pt.DataFields.Add(pt.PivotFields("amount"), "SUM of Count", SubtotalTypes.Sum)

			'Set PivotTable style
			pt.BuiltInStyle = PivotBuiltInStyles.PivotStyleMedium10

			pt.CalculateData()

			'Get slicer collection
			Dim slicers As XlsSlicerCollection = worksheet.Slicers

			Dim index As Integer = slicers.Add(pt, "E12", 0)

			Dim xlsSlicer As XlsSlicer = slicers(index)
			xlsSlicer.Name = "xlsSlicer"
			xlsSlicer.Width = 100
			xlsSlicer.Height = 120
			xlsSlicer.StyleType = SlicerStyleType.SlicerStyleLight2
			xlsSlicer.PositionLocked = True

			'Get SlicerCache object of current slicer
			Dim slicerCache As XlsSlicerCache = xlsSlicer.SlicerCache
			slicerCache.CrossFilterType = SlicerCacheCrossFilterType.ShowItemsWithNoData

			'Style setting
			Dim slicerCacheItems As XlsSlicerCacheItemCollection = xlsSlicer.SlicerCache.SlicerCacheItems
			Dim xlsSlicerCacheItem As XlsSlicerCacheItem = slicerCacheItems(0)
			xlsSlicerCacheItem.Selected = False

			Dim slicers_2 As XlsSlicerCollection = worksheet.Slicers

			Dim r1 As IPivotField = pt.PivotFields("year")
			Dim index_2 As Integer = slicers_2.Add(pt, "I12", r1)

			Dim xlsSlicer_2 As XlsSlicer = slicers(index_2)
			xlsSlicer_2.RowHeight = 40
			xlsSlicer_2.StyleType = SlicerStyleType.SlicerStyleLight3
			xlsSlicer_2.PositionLocked = False

			'Get SlicerCache object of current slicer
			Dim slicerCache_2 As XlsSlicerCache = xlsSlicer_2.SlicerCache
			slicerCache_2.CrossFilterType = SlicerCacheCrossFilterType.ShowItemsWithDataAtTop

			'Style setting
			Dim slicerCacheItems_2 As XlsSlicerCacheItemCollection = xlsSlicer_2.SlicerCache.SlicerCacheItems
			Dim xlsSlicerCacheItem_2 As XlsSlicerCacheItem = slicerCacheItems_2(1)
			xlsSlicerCacheItem_2.Selected = False
			pt.CalculateData()

			'Save to file
			wb.SaveToFile("CreateSlicerFromPivotTable.xlsx", ExcelVersion.Version2013)

			' Dispose of the workbook object to release resources
			wb.Dispose()

			' Launch the file
			ExcelDocViewer("CreateSlicerFromPivotTable.xlsx")
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
