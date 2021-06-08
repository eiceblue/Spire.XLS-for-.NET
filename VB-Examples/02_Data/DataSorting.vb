Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel

Imports Spire.Xls

Namespace DataSorting
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			Dim workbook As New Workbook()
			workbook.LoadFromFile("..\..\..\..\..\..\Data\DataSorting.xls")

			Dim worksheet As Worksheet = workbook.Worksheets(0)


			workbook.DataSorter.SortColumns.Add(2, OrderBy.Ascending)
			workbook.DataSorter.SortColumns.Add(3, OrderBy.Ascending)

			workbook.DataSorter.Sort(worksheet("A1:E19"))

			Dim result As String = "DataSorting_out.xlsx"
			workbook.SaveToFile(result,ExcelVersion.Version2013)

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

		Private Sub Form1_Load(ByVal sender As Object, ByVal e As EventArgs) Handles MyBase.Load

		End Sub

	End Class
End Namespace
