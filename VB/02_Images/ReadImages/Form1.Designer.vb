Namespace ReadImages
	Partial Public Class Form1
		Private WithEvents btnRun As Button
		Private WithEvents btnClose As Button
		Private label1 As Label
		''' <summary>
		''' Required designer variable.
		''' </summary
		Private components As System.ComponentModel.Container = Nothing



		''' <summary>
		''' Clean up any resources being used.
		''' </summary>
		Protected Overrides Sub Dispose(ByVal disposing As Boolean)
			If disposing Then
				If components IsNot Nothing Then
					components.Dispose()
				End If
			End If
			MyBase.Dispose(disposing)
		End Sub

		#Region "Windows Form Designer generated code"
		''' <summary>
		''' Required method for Designer support - do not modify
		''' the contents of this method with the code editor.
		''' </summary>
		Private Sub InitializeComponent()
			Me.btnRun = New Button()
			Me.btnClose = New Button()
			Me.label1 = New Label()
			Me.SuspendLayout()
			' 
			' btnRun
			' 
			Me.btnRun.Location = New Point(272, 56)
			Me.btnRun.Name = "btnRun"
			Me.btnRun.Size = New Size(72, 23)
			Me.btnRun.TabIndex = 2
			Me.btnRun.Text = "Read"
'			Me.btnRun.Click += New System.EventHandler(Me.btnRun_Click)
			' 
			' btnClose
			' 
			Me.btnClose.Location = New Point(352, 56)
			Me.btnClose.Name = "btnClose"
			Me.btnClose.Size = New Size(75, 23)
			Me.btnClose.TabIndex = 3
			Me.btnClose.Text = "Close"
'			Me.btnClose.Click += New System.EventHandler(Me.btnClose_Click)
			' 
			' label1
			' 
			Me.label1.AutoSize = True
			Me.label1.Font = New Font("Tahoma", 9F, FontStyle.Bold, GraphicsUnit.Point, (CByte(134)))
			Me.label1.Location = New Point(16, 16)
			Me.label1.Name = "label1"
			Me.label1.Size = New Size(406, 14)
			Me.label1.TabIndex = 4
			Me.label1.Text = "The sample demonstrates how to read images from spreadsheet."
			' 
			' Form1
			' 
			Me.AutoScaleBaseSize = New Size(6, 14)
			Me.ClientSize = New Size(456, 93)
			Me.Controls.Add(Me.label1)
			Me.Controls.Add(Me.btnClose)
			Me.Controls.Add(Me.btnRun)
			Me.MaximizeBox = False
			Me.MinimizeBox = False
			Me.Name = "Form1"
			Me.StartPosition = FormStartPosition.CenterScreen
			Me.Text = "Read Images"
			Me.ResumeLayout(False)
			Me.PerformLayout()

		End Sub
		#End Region

	End Class
End Namespace

