Imports Spire.Xls

Namespace AcceptOrRejectTrackedChanges
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			Dim workbook As New Workbook()
			workbook.LoadFromFile("..\..\..\..\..\..\Data\TrackChanges.xlsx")

			'Accept the changes or reject the changes.
			'workbook.AcceptAllTrackedChanges();
			workbook.RejectAllTrackedChanges()

			'Save to file.
			Dim outputFile As String = "AcceptOrRejectTrackedChanges.xlsx"
			workbook.SaveToFile(outputFile, FileFormat.Version2013)

			'View the document
			FileViewer(outputFile)
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
