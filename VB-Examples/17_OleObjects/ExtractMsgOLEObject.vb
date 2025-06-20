Imports Spire.Xls
Imports System.IO

Namespace ExtractMsgOLEObject
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			Dim outputFile As String = "test.msg"

			' Create a new workbook object
			Dim workbook As New Workbook()

			'Load the file from disk.
			workbook.LoadFromFile("..\..\..\..\..\..\Data\Template_Msg.xlsx")

			' Get the first worksheet
			Dim sheet As Worksheet = workbook.Worksheets(0)

			Dim type As OleObjectType
			' Determine if there is an ole object in the sheet
			If sheet.HasOleObjects Then
				For i As Integer = 0 To sheet.OleObjects.Count - 1
					Dim [Object] = sheet.OleObjects(i)
					' Get the type of ole object
					type = sheet.OleObjects(i).ObjectType
					Select Case type
						' If the type of ole object is msg
						Case OleObjectType.Msg
							File.WriteAllBytes(outputFile, [Object].OleData)
							' View the document using a file viewer
							FileViewer(outputFile)
					End Select
				Next i
			End If

			' Dispose of the workbook object
			workbook.Dispose()

			Me.Close()

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
	End Class
End Namespace
