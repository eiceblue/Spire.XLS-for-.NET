Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel

Imports Spire.Xls
Imports System.IO

Namespace ToSVG
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			Dim workbook As New Workbook()
			workbook.LoadFromFile("..\..\..\..\..\..\Data\ToSVG.xlsx")
			For i As Integer = 0 To workbook.Worksheets.Count - 1
				Dim fs As New FileStream(String.Format("sheet{0}.svg", i), FileMode.Create)
				workbook.Worksheets(i).ToSVGStream(fs, 0, 0, 0, 0)
				fs.Flush()
				fs.Close()
			Next i
			 Process.Start("sheet0.svg")
		End Sub

		Private Sub btnClose_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnClose.Click
			Close()
		End Sub
	End Class
End Namespace
