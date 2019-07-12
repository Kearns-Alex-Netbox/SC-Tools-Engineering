<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class MessageboxDifference
	Inherits System.Windows.Forms.Form

	'Form overrides dispose to clean up the component list.
	<System.Diagnostics.DebuggerNonUserCode()> _
	Protected Overrides Sub Dispose(ByVal disposing As Boolean)
		Try
			If disposing AndAlso components IsNot Nothing Then
				components.Dispose()
			End If
		Finally
			MyBase.Dispose(disposing)
		End Try
	End Sub

	'Required by the Windows Form Designer
	Private components As System.ComponentModel.IContainer

	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.  
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> _
	Private Sub InitializeComponent()
		Me.RTB_differenceReport = New System.Windows.Forms.RichTextBox()
		Me.SuspendLayout()
		'
		'RTB_differenceReport
		'
		Me.RTB_differenceReport.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
			Or System.Windows.Forms.AnchorStyles.Left) _
			Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
		Me.RTB_differenceReport.Font = New System.Drawing.Font("Consolas", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.RTB_differenceReport.Location = New System.Drawing.Point(13, 12)
		Me.RTB_differenceReport.Name = "RTB_differenceReport"
		Me.RTB_differenceReport.Size = New System.Drawing.Size(413, 293)
		Me.RTB_differenceReport.TabIndex = 1
		Me.RTB_differenceReport.Text = ""
		'
		'MessageboxDifference
		'
		Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.ClientSize = New System.Drawing.Size(438, 317)
		Me.Controls.Add(Me.RTB_differenceReport)
		Me.Name = "MessageboxDifference"
		Me.Text = "Difference"
		Me.ResumeLayout(False)

	End Sub

	Friend WithEvents RTB_differenceReport As RichTextBox
End Class
