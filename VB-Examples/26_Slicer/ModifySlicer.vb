Imports Spire.Xls
Imports Spire.Xls.Core
Imports System.ComponentModel
Imports System.Text

Namespace ModifySlicer
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

			' Get the first slicer from the slicer collection
			Dim xlsSlicer As XlsSlicer = slicers(0)

			' Set the style of the slicer to a dark theme (style type 4)
			xlsSlicer.StyleType = SlicerStyleType.SlicerStyleDark4

			' Change the caption (title) of the slicer
			xlsSlicer.Caption = "Modified Slicer"

			' Lock the position of the slicer to prevent it from being moved in the worksheet
			xlsSlicer.PositionLocked = True

			' Get the collection of cache items associated with the slicer
			Dim slicerCacheItems As XlsSlicerCacheItemCollection = xlsSlicer.SlicerCache.SlicerCacheItems

			' Get the first cache item in the collection
			Dim xlsSlicerCacheItem As XlsSlicerCacheItem = slicerCacheItems(0)

			' Deselect the cache item
			xlsSlicerCacheItem.Selected = False

			' Get the display value of the cache item
			Dim displayValue As String = xlsSlicerCacheItem.DisplayValue

			' Get the slicer cache associated with the slicer
			Dim slicerCache As XlsSlicerCache = xlsSlicer.SlicerCache

			' Set the cross-filter type to show items even if they have no associated data
			slicerCache.CrossFilterType = SlicerCacheCrossFilterType.ShowItemsWithNoData

			' Save the modified workbook to a new file with Excel 2013 version format
			wb.SaveToFile("ModifySlicer.xlsx", ExcelVersion.Version2013)

			' Dispose of the workbook object to release resources
			wb.Dispose()

			' Launch the file
			ExcelDocViewer("ModifySlicer.xlsx")
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
