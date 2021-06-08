Imports Spire.Xls
Imports Spire.Xls.Core.MergeSpreadsheet.Interfaces
Imports System.Security.Cryptography.X509Certificates

Namespace AddDigitalSignature
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a workbook
			Dim workbook As New Workbook()
			workbook.LoadFromFile("..\..\..\..\..\..\Data\DigitalSignature.xlsx")
			'Add certificate
			Dim inputFile_pfx As String = "..\..\..\..\..\..\Data\gary.pfx"
			Dim cert As New X509Certificate2(inputFile_pfx, "e-iceblue")
			'Add signature
			Dim certtime As New Date(2020, 7, 1, 7, 10, 36)
			Dim dsc As IDigitalSignatures = workbook.AddDigitalSignature(cert, "e-iceblue", certtime)

			Dim result As String = "AddDigitalSignature.xlsx"

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
	End Class
End Namespace
