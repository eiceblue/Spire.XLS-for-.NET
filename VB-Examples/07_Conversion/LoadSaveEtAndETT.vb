Imports Spire.Xls

Namespace LoadSaveEtAndETT
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new instance of Workbook
            Dim workbook As New Workbook()

            ' Load the ET file from the specified path
            workbook.LoadFromFile("..\..\..\..\..\..\Data\Sample-et.et")
            ' Uncomment the line below to load an ETT file instead
            ' workbook.LoadFromFile("..\..\..\..\..\..\Data\Sample-ett.ett");

            ' Save the workbook to an ET file with the specified output file name
            workbook.SaveToFile("result.et", FileFormat.ET)
            ' Uncomment the line below to save as ETT instead
            ' workbook.SaveToFile("result.ett", FileFormat.ETT);
            ' Release the resources used by the workbook
            workbook.Dispose()

            'view the document
            ExcelDocViewer("result.et")
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
