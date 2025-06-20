Imports Spire.Xls

Namespace AcceptOrRejectTrackedChanges
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new Workbook object
            Dim workbook As New Workbook()

            ' Load the Excel file named "TrackChanges.xlsx" from a specific path
            workbook.LoadFromFile("..\..\..\..\..\..\Data\TrackChanges.xlsx")

            ' Reject all tracked changes in the workbook
            workbook.RejectAllTrackedChanges()

            ' Specify the output file name as "AcceptOrRejectTrackedChanges.xlsx"
            Dim outputFile As String = "AcceptOrRejectTrackedChanges.xlsx"

            ' Save the workbook to the specified output file, using the Version 2013 file format
            workbook.SaveToFile(outputFile, FileFormat.Version2013)

            ' Release the resources used by the workbook
            workbook.Dispose()

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
