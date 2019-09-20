<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ImportData
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
		Me.Label3 = New System.Windows.Forms.Label()
		Me.CkB_ALPHAitems = New System.Windows.Forms.CheckBox()
		Me.RTB_Results = New System.Windows.Forms.RichTextBox()
		Me.SaveOutput_Button = New System.Windows.Forms.Button()
		Me.Import_Button = New System.Windows.Forms.Button()
		Me.Close_Button = New System.Windows.Forms.Button()
		Me.GroupBox1 = New System.Windows.Forms.GroupBox()
		Me.TB_ALPHAIndicatorLight = New System.Windows.Forms.TextBox()
		Me.TB_PCADIndicatorLight = New System.Windows.Forms.TextBox()
		Me.TB_ALPHAitemsIndicatorLight = New System.Windows.Forms.TextBox()
		Me.TB_QBIndicatorLight = New System.Windows.Forms.TextBox()
		Me.TB_QBitemsIndicatorLight = New System.Windows.Forms.TextBox()
		Me.GroupBox1.SuspendLayout
		Me.SuspendLayout
		'
		'Label3
		'
		Me.Label3.AutoSize = true
		Me.Label3.Font = New System.Drawing.Font("Consolas", 12!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Underline),System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0,Byte))
		Me.Label3.Location = New System.Drawing.Point(11, 11)
		Me.Label3.Name = "Label3"
		Me.Label3.Size = New System.Drawing.Size(108, 19)
		Me.Label3.TabIndex = 55
		Me.Label3.Text = "Import Data"
		'
		'CkB_ALPHAitems
		'
		Me.CkB_ALPHAitems.Font = New System.Drawing.Font("Microsoft Sans Serif", 12!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0,Byte))
		Me.CkB_ALPHAitems.Location = New System.Drawing.Point(15, 69)
		Me.CkB_ALPHAitems.Name = "CkB_ALPHAitems"
		Me.CkB_ALPHAitems.Size = New System.Drawing.Size(131, 44)
		Me.CkB_ALPHAitems.TabIndex = 39
		Me.CkB_ALPHAitems.Text = "Import SMT Items"
		Me.CkB_ALPHAitems.UseVisualStyleBackColor = true
		'
		'RTB_Results
		'
		Me.RTB_Results.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom)  _
            Or System.Windows.Forms.AnchorStyles.Left)  _
            Or System.Windows.Forms.AnchorStyles.Right),System.Windows.Forms.AnchorStyles)
		Me.RTB_Results.Font = New System.Drawing.Font("Consolas", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0,Byte))
		Me.RTB_Results.Location = New System.Drawing.Point(151, 14)
		Me.RTB_Results.Name = "RTB_Results"
		Me.RTB_Results.Size = New System.Drawing.Size(622, 517)
		Me.RTB_Results.TabIndex = 42
		Me.RTB_Results.Text = ""
		'
		'SaveOutput_Button
		'
		Me.SaveOutput_Button.Enabled = false
		Me.SaveOutput_Button.Font = New System.Drawing.Font("Microsoft Sans Serif", 12!)
		Me.SaveOutput_Button.Location = New System.Drawing.Point(15, 274)
		Me.SaveOutput_Button.Name = "SaveOutput_Button"
		Me.SaveOutput_Button.Size = New System.Drawing.Size(130, 30)
		Me.SaveOutput_Button.TabIndex = 40
		Me.SaveOutput_Button.Text = "Save Output"
		Me.SaveOutput_Button.UseVisualStyleBackColor = true
		'
		'Import_Button
		'
		Me.Import_Button.Font = New System.Drawing.Font("Microsoft Sans Serif", 12!)
		Me.Import_Button.Location = New System.Drawing.Point(15, 33)
		Me.Import_Button.Name = "Import_Button"
		Me.Import_Button.Size = New System.Drawing.Size(130, 30)
		Me.Import_Button.TabIndex = 38
		Me.Import_Button.Text = "Import Data"
		Me.Import_Button.UseVisualStyleBackColor = true
		'
		'Close_Button
		'
		Me.Close_Button.AutoSize = true
		Me.Close_Button.Font = New System.Drawing.Font("Microsoft Sans Serif", 12!)
		Me.Close_Button.Location = New System.Drawing.Point(15, 310)
		Me.Close_Button.Name = "Close_Button"
		Me.Close_Button.Size = New System.Drawing.Size(130, 30)
		Me.Close_Button.TabIndex = 41
		Me.Close_Button.Text = "Close"
		Me.Close_Button.UseVisualStyleBackColor = true
		'
		'GroupBox1
		'
		Me.GroupBox1.BackColor = System.Drawing.Color.Transparent
		Me.GroupBox1.Controls.Add(Me.TB_ALPHAIndicatorLight)
		Me.GroupBox1.Controls.Add(Me.TB_PCADIndicatorLight)
		Me.GroupBox1.Controls.Add(Me.TB_ALPHAitemsIndicatorLight)
		Me.GroupBox1.Controls.Add(Me.TB_QBIndicatorLight)
		Me.GroupBox1.Controls.Add(Me.TB_QBitemsIndicatorLight)
		Me.GroupBox1.Location = New System.Drawing.Point(15, 119)
		Me.GroupBox1.Name = "GroupBox1"
		Me.GroupBox1.Size = New System.Drawing.Size(130, 149)
		Me.GroupBox1.TabIndex = 57
		Me.GroupBox1.TabStop = false
		Me.GroupBox1.Text = "Results"
		'
		'TB_ALPHAIndicatorLight
		'
		Me.TB_ALPHAIndicatorLight.BackColor = System.Drawing.Color.White
		Me.TB_ALPHAIndicatorLight.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me.TB_ALPHAIndicatorLight.Cursor = System.Windows.Forms.Cursors.Hand
		Me.TB_ALPHAIndicatorLight.ForeColor = System.Drawing.Color.Black
		Me.TB_ALPHAIndicatorLight.Location = New System.Drawing.Point(6, 19)
		Me.TB_ALPHAIndicatorLight.Name = "TB_ALPHAIndicatorLight"
		Me.TB_ALPHAIndicatorLight.ReadOnly = true
		Me.TB_ALPHAIndicatorLight.Size = New System.Drawing.Size(118, 20)
		Me.TB_ALPHAIndicatorLight.TabIndex = 43
		Me.TB_ALPHAIndicatorLight.Text = "SMT BOM"
		Me.TB_ALPHAIndicatorLight.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
		'
		'TB_PCADIndicatorLight
		'
		Me.TB_PCADIndicatorLight.BackColor = System.Drawing.Color.White
		Me.TB_PCADIndicatorLight.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me.TB_PCADIndicatorLight.Cursor = System.Windows.Forms.Cursors.Hand
		Me.TB_PCADIndicatorLight.ForeColor = System.Drawing.Color.Black
		Me.TB_PCADIndicatorLight.Location = New System.Drawing.Point(6, 71)
		Me.TB_PCADIndicatorLight.Name = "TB_PCADIndicatorLight"
		Me.TB_PCADIndicatorLight.ReadOnly = true
		Me.TB_PCADIndicatorLight.Size = New System.Drawing.Size(118, 20)
		Me.TB_PCADIndicatorLight.TabIndex = 45
		Me.TB_PCADIndicatorLight.Text = "PCAD BOM"
		Me.TB_PCADIndicatorLight.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
		'
		'TB_ALPHAitemsIndicatorLight
		'
		Me.TB_ALPHAitemsIndicatorLight.BackColor = System.Drawing.Color.White
		Me.TB_ALPHAitemsIndicatorLight.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me.TB_ALPHAitemsIndicatorLight.Cursor = System.Windows.Forms.Cursors.Hand
		Me.TB_ALPHAitemsIndicatorLight.ForeColor = System.Drawing.Color.Black
		Me.TB_ALPHAitemsIndicatorLight.Location = New System.Drawing.Point(6, 45)
		Me.TB_ALPHAitemsIndicatorLight.Name = "TB_ALPHAitemsIndicatorLight"
		Me.TB_ALPHAitemsIndicatorLight.ReadOnly = true
		Me.TB_ALPHAitemsIndicatorLight.Size = New System.Drawing.Size(118, 20)
		Me.TB_ALPHAitemsIndicatorLight.TabIndex = 44
		Me.TB_ALPHAitemsIndicatorLight.Text = "SMT Items"
		Me.TB_ALPHAitemsIndicatorLight.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
		'
		'TB_QBIndicatorLight
		'
		Me.TB_QBIndicatorLight.BackColor = System.Drawing.Color.White
		Me.TB_QBIndicatorLight.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me.TB_QBIndicatorLight.Cursor = System.Windows.Forms.Cursors.Hand
		Me.TB_QBIndicatorLight.ForeColor = System.Drawing.Color.Black
		Me.TB_QBIndicatorLight.Location = New System.Drawing.Point(6, 123)
		Me.TB_QBIndicatorLight.Name = "TB_QBIndicatorLight"
		Me.TB_QBIndicatorLight.ReadOnly = true
		Me.TB_QBIndicatorLight.Size = New System.Drawing.Size(118, 20)
		Me.TB_QBIndicatorLight.TabIndex = 46
		Me.TB_QBIndicatorLight.Text = "QB BOM"
		Me.TB_QBIndicatorLight.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
		'
		'TB_QBitemsIndicatorLight
		'
		Me.TB_QBitemsIndicatorLight.BackColor = System.Drawing.Color.White
		Me.TB_QBitemsIndicatorLight.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me.TB_QBitemsIndicatorLight.Cursor = System.Windows.Forms.Cursors.Hand
		Me.TB_QBitemsIndicatorLight.ForeColor = System.Drawing.Color.Black
		Me.TB_QBitemsIndicatorLight.Location = New System.Drawing.Point(6, 97)
		Me.TB_QBitemsIndicatorLight.Name = "TB_QBitemsIndicatorLight"
		Me.TB_QBitemsIndicatorLight.ReadOnly = true
		Me.TB_QBitemsIndicatorLight.Size = New System.Drawing.Size(118, 20)
		Me.TB_QBitemsIndicatorLight.TabIndex = 48
		Me.TB_QBitemsIndicatorLight.Text = "QB Items"
		Me.TB_QBitemsIndicatorLight.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
		'
		'ImportData
		'
		Me.AutoScaleDimensions = New System.Drawing.SizeF(6!, 13!)
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.AutoScroll = true
		Me.ClientSize = New System.Drawing.Size(784, 544)
		Me.Controls.Add(Me.GroupBox1)
		Me.Controls.Add(Me.Label3)
		Me.Controls.Add(Me.CkB_ALPHAitems)
		Me.Controls.Add(Me.RTB_Results)
		Me.Controls.Add(Me.SaveOutput_Button)
		Me.Controls.Add(Me.Import_Button)
		Me.Controls.Add(Me.Close_Button)
		Me.Name = "ImportData"
		Me.Text = "Import Data"
		Me.GroupBox1.ResumeLayout(false)
		Me.GroupBox1.PerformLayout
		Me.ResumeLayout(false)
		Me.PerformLayout

End Sub

	Friend WithEvents Label3 As Label
	Friend WithEvents CkB_ALPHAitems As CheckBox
	Friend WithEvents RTB_Results As RichTextBox
	Friend WithEvents SaveOutput_Button As Button
	Friend WithEvents Import_Button As Button
	Friend WithEvents Close_Button As Button
	Friend WithEvents GroupBox1 As GroupBox
	Friend WithEvents TB_ALPHAIndicatorLight As TextBox
	Friend WithEvents TB_PCADIndicatorLight As TextBox
	Friend WithEvents TB_ALPHAitemsIndicatorLight As TextBox
	Friend WithEvents TB_QBIndicatorLight As TextBox
	Friend WithEvents TB_QBitemsIndicatorLight As TextBox
End Class
