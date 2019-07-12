<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class MessageBoxOption
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
		Me.B_OK = New System.Windows.Forms.Button()
		Me.B_Cancel = New System.Windows.Forms.Button()
		Me.Label1 = New System.Windows.Forms.Label()
		Me.TB_Options = New System.Windows.Forms.TextBox()
		Me.SuspendLayout()
		'
		'B_OK
		'
		Me.B_OK.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
		Me.B_OK.AutoSize = True
		Me.B_OK.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
		Me.B_OK.Location = New System.Drawing.Point(111, 64)
		Me.B_OK.Name = "B_OK"
		Me.B_OK.Size = New System.Drawing.Size(32, 23)
		Me.B_OK.TabIndex = 0
		Me.B_OK.Text = "OK"
		Me.B_OK.UseVisualStyleBackColor = True
		'
		'B_Cancel
		'
		Me.B_Cancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
		Me.B_Cancel.AutoSize = True
		Me.B_Cancel.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
		Me.B_Cancel.Location = New System.Drawing.Point(149, 64)
		Me.B_Cancel.Name = "B_Cancel"
		Me.B_Cancel.Size = New System.Drawing.Size(50, 23)
		Me.B_Cancel.TabIndex = 1
		Me.B_Cancel.Text = "Cancel"
		Me.B_Cancel.UseVisualStyleBackColor = True
		'
		'Label1
		'
		Me.Label1.AutoSize = True
		Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label1.Location = New System.Drawing.Point(13, 13)
		Me.Label1.Name = "Label1"
		Me.Label1.Size = New System.Drawing.Size(191, 18)
		Me.Label1.TabIndex = 2
		Me.Label1.Text = "Enter Options Alphabetically"
		'
		'TB_Options
		'
		Me.TB_Options.Location = New System.Drawing.Point(16, 34)
		Me.TB_Options.Name = "TB_Options"
		Me.TB_Options.Size = New System.Drawing.Size(183, 20)
		Me.TB_Options.TabIndex = 3
		'
		'MessageBoxOption
		'
		Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.ClientSize = New System.Drawing.Size(211, 99)
		Me.Controls.Add(Me.TB_Options)
		Me.Controls.Add(Me.Label1)
		Me.Controls.Add(Me.B_Cancel)
		Me.Controls.Add(Me.B_OK)
		Me.Name = "MessageBoxOption"
		Me.Text = "MessageBoxOption"
		Me.ResumeLayout(False)
		Me.PerformLayout()

	End Sub

	Friend WithEvents B_OK As Button
	Friend WithEvents B_Cancel As Button
	Friend WithEvents Label1 As Label
	Friend WithEvents TB_Options As TextBox
End Class
