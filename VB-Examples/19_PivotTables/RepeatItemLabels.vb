Imports Spire.Xls
Imports System.ComponentModel
Imports System.Text
Imports Spire.Xls.Core.Spreadsheet.PivotTables

Namespace RepeatItemLabels
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a workbook
			Dim workbook As New Workbook()
			'Load an excel file including pivot table
			workbook.LoadFromFile("..\..\..\..\..\..\Data\RepeatItemLabelsExample.xlsx")
			'Get the first worksheet
			Dim sheet As Worksheet = workbook.Worksheets(0)
			'Add an empty worksheet 
			Dim sheet2 As Worksheet = workbook.CreateEmptySheet()
			'Add PivotTable
			sheet2.Name = "Pivot Table"
			Dim dataRange As CellRange = sheet.Range("A1:D9")
			Dim cache As PivotCache = workbook.PivotCaches.Add(dataRange)
			Dim pt As PivotTable = sheet2.PivotTables.Add("Pivot Table", sheet.Range("A1"), cache)
			Dim r1 = pt.PivotFields("VendorNo")
			r1.Axis = AxisTypes.Row
			pt.Options.RowHeaderCaption = "VendorNo"
			r1.Subtotals = SubtotalTypes.None

			r1.RepeatItemLabels = True
			'Repeat item lables
			pt.PivotFields("OnHand").RepeatItemLabels = True
			pt.Options.RowLayout = PivotTableLayoutType.Tabular
			Dim r2 = pt.PivotFields("Desc")
			r2.Axis = AxisTypes.Row
			pt.DataFields.Add(pt.PivotFields("OnHand"), "Sum of onHand", SubtotalTypes.None)
			pt.BuiltInStyle = PivotBuiltInStyles.PivotStyleMedium12
			Dim result As String = "RepeatItemLabels_result.xlsx"
			'Save to file
			workbook.SaveToFile(result, ExcelVersion.Version2010)

			'View the document
			FileViewer(result)
		End Sub

		Private Sub FileViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub

		Private Sub btnClose_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnClose.Click
			Close()
		End Sub

		Private Sub Form1_Load(ByVal sender As Object, ByVal e As EventArgs) Handles MyBase.Load

		End Sub
	End Class
End Namespace
