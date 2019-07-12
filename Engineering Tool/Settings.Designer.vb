<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Settings
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
		Me.Remind_CheckBox = New System.Windows.Forms.CheckBox()
		Me.Label2 = New System.Windows.Forms.Label()
		Me.DatabaseFile_TextBox = New System.Windows.Forms.TextBox()
		Me.BrowseDatabaseFile_Button = New System.Windows.Forms.Button()
		Me.L_Version = New System.Windows.Forms.Label()
		Me.Label1 = New System.Windows.Forms.Label()
		Me.TB_ProjectDirectory = New System.Windows.Forms.TextBox()
		Me.BrowseReleaseLocation_Button = New System.Windows.Forms.Button()
		Me.B_Save = New System.Windows.Forms.Button()
		Me.B_Cancel = New System.Windows.Forms.Button()
		Me.TestConnection_Button = New System.Windows.Forms.Button()
		Me.Label3 = New System.Windows.Forms.Label()
		Me.RevisionExcel_TextBox = New System.Windows.Forms.TextBox()
		Me.RevisionBrowse_Button = New System.Windows.Forms.Button()
		Me.SuspendLayout()
		'
		'Remind_CheckBox
		'
		Me.Remind_CheckBox.AutoSize = True
		Me.Remind_CheckBox.Checked = True
		Me.Remind_CheckBox.CheckState = System.Windows.Forms.CheckState.Checked
		Me.Remind_CheckBox.Location = New System.Drawing.Point(125, 98)
		Me.Remind_CheckBox.Name = "Remind_CheckBox"
		Me.Remind_CheckBox.Size = New System.Drawing.Size(79, 17)
		Me.Remind_CheckBox.TabIndex = 42
		Me.Remind_CheckBox.Text = "Remind me"
		Me.Remind_CheckBox.UseVisualStyleBackColor = True
		'
		'Label2
		'
		Me.Label2.AutoSize = True
		Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label2.Location = New System.Drawing.Point(11, 96)
		Me.Label2.Name = "Label2"
		Me.Label2.Size = New System.Drawing.Size(108, 20)
		Me.Label2.TabIndex = 39
		Me.Label2.Text = "Database File"
		'
		'DatabaseFile_TextBox
		'
		Me.DatabaseFile_TextBox.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.DatabaseFile_TextBox.Location = New System.Drawing.Point(11, 119)
		Me.DatabaseFile_TextBox.Name = "DatabaseFile_TextBox"
		Me.DatabaseFile_TextBox.Size = New System.Drawing.Size(259, 26)
		Me.DatabaseFile_TextBox.TabIndex = 40
		'
		'BrowseDatabaseFile_Button
		'
		Me.BrowseDatabaseFile_Button.AutoSize = True
		Me.BrowseDatabaseFile_Button.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
		Me.BrowseDatabaseFile_Button.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!)
		Me.BrowseDatabaseFile_Button.Location = New System.Drawing.Point(11, 151)
		Me.BrowseDatabaseFile_Button.Name = "BrowseDatabaseFile_Button"
		Me.BrowseDatabaseFile_Button.Size = New System.Drawing.Size(72, 30)
		Me.BrowseDatabaseFile_Button.TabIndex = 41
		Me.BrowseDatabaseFile_Button.Text = "Browse"
		Me.BrowseDatabaseFile_Button.UseVisualStyleBackColor = True
		'
		'L_Version
		'
		Me.L_Version.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
		Me.L_Version.AutoSize = True
		Me.L_Version.Location = New System.Drawing.Point(11, 291)
		Me.L_Version.Name = "L_Version"
		Me.L_Version.Size = New System.Drawing.Size(53, 13)
		Me.L_Version.TabIndex = 38
		Me.L_Version.Text = "V: 0.0.0.0"
		'
		'Label1
		'
		Me.Label1.AutoSize = True
		Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label1.Location = New System.Drawing.Point(11, 8)
		Me.Label1.Name = "Label1"
		Me.Label1.Size = New System.Drawing.Size(125, 20)
		Me.Label1.TabIndex = 33
		Me.Label1.Text = "Project Directory"
		'
		'TB_ProjectDirectory
		'
		Me.TB_ProjectDirectory.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.TB_ProjectDirectory.Location = New System.Drawing.Point(11, 31)
		Me.TB_ProjectDirectory.Name = "TB_ProjectDirectory"
		Me.TB_ProjectDirectory.Size = New System.Drawing.Size(259, 26)
		Me.TB_ProjectDirectory.TabIndex = 34
		'
		'BrowseReleaseLocation_Button
		'
		Me.BrowseReleaseLocation_Button.AutoSize = True
		Me.BrowseReleaseLocation_Button.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
		Me.BrowseReleaseLocation_Button.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!)
		Me.BrowseReleaseLocation_Button.Location = New System.Drawing.Point(11, 63)
		Me.BrowseReleaseLocation_Button.Name = "BrowseReleaseLocation_Button"
		Me.BrowseReleaseLocation_Button.Size = New System.Drawing.Size(72, 30)
		Me.BrowseReleaseLocation_Button.TabIndex = 35
		Me.BrowseReleaseLocation_Button.Text = "Browse"
		Me.BrowseReleaseLocation_Button.UseVisualStyleBackColor = True
		'
		'B_Save
		'
		Me.B_Save.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
		Me.B_Save.AutoSize = True
		Me.B_Save.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
		Me.B_Save.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!)
		Me.B_Save.Location = New System.Drawing.Point(139, 276)
		Me.B_Save.Name = "B_Save"
		Me.B_Save.Size = New System.Drawing.Size(55, 30)
		Me.B_Save.TabIndex = 36
		Me.B_Save.Text = "Save"
		Me.B_Save.UseVisualStyleBackColor = True
		'
		'B_Cancel
		'
		Me.B_Cancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
		Me.B_Cancel.AutoSize = True
		Me.B_Cancel.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
		Me.B_Cancel.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!)
		Me.B_Cancel.Location = New System.Drawing.Point(200, 276)
		Me.B_Cancel.Name = "B_Cancel"
		Me.B_Cancel.Size = New System.Drawing.Size(68, 30)
		Me.B_Cancel.TabIndex = 37
		Me.B_Cancel.Text = "Cancel"
		Me.B_Cancel.UseVisualStyleBackColor = True
		'
		'TestConnection_Button
		'
		Me.TestConnection_Button.AutoSize = True
		Me.TestConnection_Button.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
		Me.TestConnection_Button.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!)
		Me.TestConnection_Button.Location = New System.Drawing.Point(135, 151)
		Me.TestConnection_Button.Name = "TestConnection_Button"
		Me.TestConnection_Button.Size = New System.Drawing.Size(135, 30)
		Me.TestConnection_Button.TabIndex = 43
		Me.TestConnection_Button.Text = "Test Connection"
		Me.TestConnection_Button.UseVisualStyleBackColor = True
		'
		'Label3
		'
		Me.Label3.AutoSize = True
		Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label3.Location = New System.Drawing.Point(11, 184)
		Me.Label3.Name = "Label3"
		Me.Label3.Size = New System.Drawing.Size(140, 20)
		Me.Label3.TabIndex = 44
		Me.Label3.Text = "Revision Excel File"
		'
		'RevisionExcel_TextBox
		'
		Me.RevisionExcel_TextBox.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.RevisionExcel_TextBox.Location = New System.Drawing.Point(11, 207)
		Me.RevisionExcel_TextBox.Name = "RevisionExcel_TextBox"
		Me.RevisionExcel_TextBox.Size = New System.Drawing.Size(259, 26)
		Me.RevisionExcel_TextBox.TabIndex = 45
		'
		'RevisionBrowse_Button
		'
		Me.RevisionBrowse_Button.AutoSize = True
		Me.RevisionBrowse_Button.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
		Me.RevisionBrowse_Button.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!)
		Me.RevisionBrowse_Button.Location = New System.Drawing.Point(11, 239)
		Me.RevisionBrowse_Button.Name = "RevisionBrowse_Button"
		Me.RevisionBrowse_Button.Size = New System.Drawing.Size(72, 30)
		Me.RevisionBrowse_Button.TabIndex = 46
		Me.RevisionBrowse_Button.Text = "Browse"
		Me.RevisionBrowse_Button.UseVisualStyleBackColor = True
		'
		'Settings
		'
		Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.ClientSize = New System.Drawing.Size(281, 314)
		Me.Controls.Add(Me.Label3)
		Me.Controls.Add(Me.RevisionExcel_TextBox)
		Me.Controls.Add(Me.RevisionBrowse_Button)
		Me.Controls.Add(Me.TestConnection_Button)
		Me.Controls.Add(Me.Remind_CheckBox)
		Me.Controls.Add(Me.Label2)
		Me.Controls.Add(Me.DatabaseFile_TextBox)
		Me.Controls.Add(Me.BrowseDatabaseFile_Button)
		Me.Controls.Add(Me.L_Version)
		Me.Controls.Add(Me.Label1)
		Me.Controls.Add(Me.TB_ProjectDirectory)
		Me.Controls.Add(Me.BrowseReleaseLocation_Button)
		Me.Controls.Add(Me.B_Save)
		Me.Controls.Add(Me.B_Cancel)
		Me.Name = "Settings"
		Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
		Me.Text = "Settings"
		Me.ResumeLayout(False)
		Me.PerformLayout()

	End Sub

	Friend WithEvents Remind_CheckBox As CheckBox
	Friend WithEvents Label2 As Label
	Friend WithEvents DatabaseFile_TextBox As TextBox
	Friend WithEvents BrowseDatabaseFile_Button As Button
	Friend WithEvents L_Version As Label
	Friend WithEvents Label1 As Label
	Friend WithEvents TB_ProjectDirectory As TextBox
	Friend WithEvents BrowseReleaseLocation_Button As Button
	Friend WithEvents B_Save As Button
	Friend WithEvents B_Cancel As Button
	Friend WithEvents TestConnection_Button As Button
	Friend WithEvents Label3 As Label
	Friend WithEvents RevisionExcel_TextBox As TextBox
	Friend WithEvents RevisionBrowse_Button As Button
End Class
