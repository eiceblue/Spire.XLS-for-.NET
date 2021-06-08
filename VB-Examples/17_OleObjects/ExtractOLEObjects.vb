Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel

Imports Spire.Xls
Imports System.IO

Namespace ExtractOLEObjects
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a workbook
			Dim workbook As New Workbook()

			'Load the document from disk
			workbook.LoadFromFile("..\..\..\..\..\..\Data\ExtractOle2.xlsx")

			'Get the first worksheet
			Dim sheet As Worksheet = workbook.Worksheets(0)

			'Extract ole objects
			If sheet.HasOleObjects Then
				For i As Integer = 0 To sheet.OleObjects.Count - 1
					Dim [Object] = sheet.OleObjects(i)
					Dim type As OleObjectType = sheet.OleObjects(i).ObjectType
					Select Case type
						'Word document
						Case OleObjectType.WordDocument
							File.WriteAllBytes("Ole.docx", [Object].OleData)
						'PowerPoint document
						Case OleObjectType.PowerPointSlide
							File.WriteAllBytes("Ole.pptx", [Object].OleData)
						'PDF document
						Case OleObjectType.AdobeAcrobatDocument
							File.WriteAllBytes("Ole.pdf", [Object].OleData)
					End Select
				Next i
			End If
			MessageBox.Show("Completed!")
		End Sub

		Private Sub btnClose_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnClose.Click
			Close()
		End Sub
	End Class
End Namespace
