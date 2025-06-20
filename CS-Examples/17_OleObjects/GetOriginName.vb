Imports Spire.Xls
Imports System.IO

Namespace GetOriginName
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			' Load an existing workbook from a file
			Dim workbook As New Workbook()
			workbook.LoadFromFile("..\..\..\..\..\..\Data\GetOriginName.xlsx")

			' Get the first sheet
			Dim worksheet As Worksheet = workbook.Worksheets(0)

			Dim information As String = ""
			' Check if the worksheet contains any OLE objects
			If worksheet.HasOleObjects Then
				' Iterate over each OLE object in the worksheet
				For i As Integer = 0 To worksheet.OleObjects.Count - 1
					' Get the current OLE object
					Dim [Object] = worksheet.OleObjects(i)

					' Determine the type of the OLE object
					Dim type As OleObjectType = worksheet.OleObjects(i).ObjectType
					information &= "Type: " & type.ToString() & vbLf

					' Determine the origin name of the OLE object
					Dim originName As String = worksheet.OleObjects(i).OleOriginName
					information &= "Origin Name: " & originName & vbLf
				Next i
			End If

			' Dispose of the workbook object to release resources
			workbook.Dispose()

			' Save the information to a TXT file
			Dim result As String = "GetOriginName-out.txt"
			File.WriteAllText(result,information)

			' Launch the file
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
	End Class


End Namespace
