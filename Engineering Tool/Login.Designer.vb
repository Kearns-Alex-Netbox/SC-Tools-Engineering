<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Login
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
		Me.Database_Label = New System.Windows.Forms.Label()
		Me.Label3 = New System.Windows.Forms.Label()
		Me.L_Version = New System.Windows.Forms.Label()
		Me.B_Exit = New System.Windows.Forms.Button()
		Me.TB_Password = New System.Windows.Forms.TextBox()
		Me.Label2 = New System.Windows.Forms.Label()
		Me.TB_User = New System.Windows.Forms.TextBox()
		Me.Label1 = New System.Windows.Forms.Label()
		Me.B_Login = New System.Windows.Forms.Button()
		Me.SuspendLayout()
		'
		'Database_Label
		'
		Me.Database_Label.AutoSize = True
		Me.Database_Label.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Database_Label.Location = New System.Drawing.Point(10, 105)
		Me.Database_Label.Name = "Database_Label"
		Me.Database_Label.Size = New System.Drawing.Size(112, 25)
		Me.Database_Label.TabIndex = 35
		Me.Database_Label.Text = "Database"
		'
		'Label3
		'
		Me.Label3.AutoSize = True
		Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 18.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label3.Location = New System.Drawing.Point(31, 8)
		Me.Label3.Name = "Label3"
		Me.Label3.Size = New System.Drawing.Size(216, 29)
		Me.Label3.TabIndex = 27
		Me.Label3.Text = "Engineering Tool"
		'
		'L_Version
		'
		Me.L_Version.AutoSize = True
		Me.L_Version.Location = New System.Drawing.Point(12, 137)
		Me.L_Version.Name = "L_Version"
		Me.L_Version.Size = New System.Drawing.Size(53, 13)
		Me.L_Version.TabIndex = 34
		Me.L_Version.Text = "V: 0.0.0.0"
		'
		'B_Exit
		'
		Me.B_Exit.AutoSize = True
		Me.B_Exit.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
		Me.B_Exit.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.B_Exit.Location = New System.Drawing.Point(220, 104)
		Me.B_Exit.Name = "B_Exit"
		Me.B_Exit.Size = New System.Drawing.Size(49, 30)
		Me.B_Exit.TabIndex = 33
		Me.B_Exit.Text = "Exit"
		Me.B_Exit.UseVisualStyleBackColor = True
		'
		'TB_Password
		'
		Me.TB_Password.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.TB_Password.Location = New System.Drawing.Point(109, 72)
		Me.TB_Password.Name = "TB_Password"
		Me.TB_Password.Size = New System.Drawing.Size(160, 26)
		Me.TB_Password.TabIndex = 31
		Me.TB_Password.UseSystemPasswordChar = True
		'
		'Label2
		'
		Me.Label2.AutoSize = True
		Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label2.Location = New System.Drawing.Point(11, 75)
		Me.Label2.Name = "Label2"
		Me.Label2.Size = New System.Drawing.Size(82, 20)
		Me.Label2.TabIndex = 30
		Me.Label2.Text = "Password:"
		'
		'TB_User
		'
		Me.TB_User.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.TB_User.Location = New System.Drawing.Point(109, 40)
		Me.TB_User.Name = "TB_User"
		Me.TB_User.Size = New System.Drawing.Size(160, 26)
		Me.TB_User.TabIndex = 29
		'
		'Label1
		'
		Me.Label1.AutoSize = True
		Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label1.Location = New System.Drawing.Point(11, 43)
		Me.Label1.Name = "Label1"
		Me.Label1.Size = New System.Drawing.Size(93, 20)
		Me.Label1.TabIndex = 28
		Me.Label1.Text = "User Name:"
		'
		'B_Login
		'
		Me.B_Login.AutoSize = True
		Me.B_Login.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
		Me.B_Login.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.B_Login.Location = New System.Drawing.Point(151, 104)
		Me.B_Login.Name = "B_Login"
		Me.B_Login.Size = New System.Drawing.Size(63, 30)
		Me.B_Login.TabIndex = 32
		Me.B_Login.Text = "Login"
		Me.B_Login.UseVisualStyleBackColor = True
		'
		'Form1
		'
		Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.ClientSize = New System.Drawing.Size(279, 159)
		Me.Controls.Add(Me.Database_Label)
		Me.Controls.Add(Me.Label3)
		Me.Controls.Add(Me.L_Version)
		Me.Controls.Add(Me.B_Exit)
		Me.Controls.Add(Me.TB_Password)
		Me.Controls.Add(Me.Label2)
		Me.Controls.Add(Me.TB_User)
		Me.Controls.Add(Me.Label1)
		Me.Controls.Add(Me.B_Login)
		Me.Name = "Form1"
		Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
		Me.Text = "Login"
		Me.ResumeLayout(False)
		Me.PerformLayout()

	End Sub

	Friend WithEvents Database_Label As Label
	Friend WithEvents Label3 As Label
	Friend WithEvents L_Version As Label
	Friend WithEvents B_Exit As Button
	Friend WithEvents TB_Password As TextBox
	Friend WithEvents Label2 As Label
	Friend WithEvents TB_User As TextBox
	Friend WithEvents Label1 As Label
	Friend WithEvents B_Login As Button
End Class
