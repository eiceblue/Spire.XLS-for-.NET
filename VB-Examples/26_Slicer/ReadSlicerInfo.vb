Imports Spire.Xls
Imports Spire.Xls.Core
Imports System.ComponentModel
Imports System.Text
Imports System.IO

Namespace ReadSlicerInfo
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

			Dim builder As New StringBuilder()

			builder.AppendLine("slicers.Count：" & slicers.Count)

			For i As Integer = 0 To slicers.Count - 1
				Dim xlsSlicer As XlsSlicer = slicers(i)
				builder.AppendLine()
				builder.AppendLine("xlsSlicer.Name：" & xlsSlicer.Name)
				builder.AppendLine("xlsSlicer.Caption：" & xlsSlicer.Caption)
				builder.AppendLine("xlsSlicer.NumberOfColumns：" & xlsSlicer.NumberOfColumns)
				builder.AppendLine("xlsSlicer.ColumnWidth：" & xlsSlicer.ColumnWidth)
				builder.AppendLine("xlsSlicer.RowHeight：" & xlsSlicer.RowHeight)
				builder.AppendLine("xlsSlicer.ShowCaption：" & xlsSlicer.ShowCaption)
				builder.AppendLine("xlsSlicer.PositionLocked：" & xlsSlicer.PositionLocked)
				builder.AppendLine("xlsSlicer.Width：" & xlsSlicer.Width)
				builder.AppendLine("xlsSlicer.Height：" & xlsSlicer.Height)

				Dim slicerCache As XlsSlicerCache = xlsSlicer.SlicerCache

				builder.AppendLine("slicerCache.SourceName：" & slicerCache.SourceName)
				builder.AppendLine("slicerCache.IsTabular：" & slicerCache.IsTabular)
				builder.AppendLine("slicerCache.Name：" & slicerCache.Name)

				Dim slicerCacheItems As XlsSlicerCacheItemCollection = slicerCache.SlicerCacheItems
				Dim xlsSlicerCacheItem As XlsSlicerCacheItem = slicerCacheItems(1)

				builder.AppendLine("xlsSlicerCacheItem.Selected：" & xlsSlicerCacheItem.Selected)
			Next i

			File.WriteAllText("ReadSlicerInfo.txt", builder.ToString())

			' Dispose of the workbook object to release resources
			wb.Dispose()

			' Launch the file
			ExcelDocViewer("ReadSlicerInfo.txt")
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
