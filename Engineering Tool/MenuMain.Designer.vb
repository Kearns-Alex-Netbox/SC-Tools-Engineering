<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class MenuMain
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
		Me.P_Menu = New System.Windows.Forms.Panel()
		Me.Label2 = New System.Windows.Forms.Label()
		Me.L_LastALPHAImport = New System.Windows.Forms.Label()
		Me.Label1 = New System.Windows.Forms.Label()
		Me.L_LastImport = New System.Windows.Forms.Label()
		Me.B_BoardOption = New System.Windows.Forms.Button()
		Me.B_PartUsage = New System.Windows.Forms.Button()
		Me.ImportData_Button = New System.Windows.Forms.Button()
		Me.B_FindStockNumber = New System.Windows.Forms.Button()
		Me.B_MasterWindow = New System.Windows.Forms.Button()
		Me.B_Settings = New System.Windows.Forms.Button()
		Me.B_CompareRevisions = New System.Windows.Forms.Button()
		Me.B_Close = New System.Windows.Forms.Button()
		Me.B_Cascade = New System.Windows.Forms.Button()
		Me.B_Minimize = New System.Windows.Forms.Button()
		Me.B_Tile = New System.Windows.Forms.Button()
		Me.B_Exit = New System.Windows.Forms.Button()
		Me.GroupBox1 = New System.Windows.Forms.GroupBox()
		Me.Panel1 = New System.Windows.Forms.Panel()
		Me.P_Menu.SuspendLayout()
		Me.GroupBox1.SuspendLayout()
		Me.Panel1.SuspendLayout()
		Me.SuspendLayout()
		'
		'P_Menu
		'
		Me.P_Menu.AutoScroll = True
		Me.P_Menu.BackColor = System.Drawing.Color.Silver
		Me.P_Menu.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.P_Menu.Controls.Add(Me.Label2)
		Me.P_Menu.Controls.Add(Me.L_LastALPHAImport)
		Me.P_Menu.Controls.Add(Me.Label1)
		Me.P_Menu.Controls.Add(Me.L_LastImport)
		Me.P_Menu.Controls.Add(Me.B_BoardOption)
		Me.P_Menu.Controls.Add(Me.B_PartUsage)
		Me.P_Menu.Controls.Add(Me.ImportData_Button)
		Me.P_Menu.Controls.Add(Me.B_FindStockNumber)
		Me.P_Menu.Controls.Add(Me.B_MasterWindow)
		Me.P_Menu.Controls.Add(Me.B_Settings)
		Me.P_Menu.Controls.Add(Me.B_CompareRevisions)
		Me.P_Menu.Dock = System.Windows.Forms.DockStyle.Left
		Me.P_Menu.Location = New System.Drawing.Point(0, 0)
		Me.P_Menu.Name = "P_Menu"
		Me.P_Menu.Size = New System.Drawing.Size(208, 673)
		Me.P_Menu.TabIndex = 7
		'
		'Label2
		'
		Me.Label2.Anchor = System.Windows.Forms.AnchorStyles.Top
		Me.Label2.AutoSize = True
		Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label2.Location = New System.Drawing.Point(29, 52)
		Me.Label2.Name = "Label2"
		Me.Label2.Size = New System.Drawing.Size(147, 20)
		Me.Label2.TabIndex = 2
		Me.Label2.Text = "Last ALPHA Import"
		Me.Label2.TextAlign = System.Drawing.ContentAlignment.TopCenter
		'
		'L_LastALPHAImport
		'
		Me.L_LastALPHAImport.Anchor = System.Windows.Forms.AnchorStyles.Top
		Me.L_LastALPHAImport.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.L_LastALPHAImport.Location = New System.Drawing.Point(0, 72)
		Me.L_LastALPHAImport.Name = "L_LastALPHAImport"
		Me.L_LastALPHAImport.Size = New System.Drawing.Size(201, 20)
		Me.L_LastALPHAImport.TabIndex = 3
		Me.L_LastALPHAImport.Text = "Date and Time"
		Me.L_LastALPHAImport.TextAlign = System.Drawing.ContentAlignment.TopCenter
		'
		'Label1
		'
		Me.Label1.Anchor = System.Windows.Forms.AnchorStyles.Top
		Me.Label1.AutoSize = True
		Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label1.Location = New System.Drawing.Point(57, 7)
		Me.Label1.Name = "Label1"
		Me.Label1.Size = New System.Drawing.Size(90, 20)
		Me.Label1.TabIndex = 0
		Me.Label1.Text = "Last Import"
		Me.Label1.TextAlign = System.Drawing.ContentAlignment.TopCenter
		'
		'L_LastImport
		'
		Me.L_LastImport.Anchor = System.Windows.Forms.AnchorStyles.Top
		Me.L_LastImport.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.L_LastImport.Location = New System.Drawing.Point(0, 27)
		Me.L_LastImport.Name = "L_LastImport"
		Me.L_LastImport.Size = New System.Drawing.Size(201, 20)
		Me.L_LastImport.TabIndex = 1
		Me.L_LastImport.Text = "Date and Time"
		Me.L_LastImport.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		'
		'B_BoardOption
		'
		Me.B_BoardOption.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
		Me.B_BoardOption.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!)
		Me.B_BoardOption.Location = New System.Drawing.Point(20, 200)
		Me.B_BoardOption.Name = "B_BoardOption"
		Me.B_BoardOption.Size = New System.Drawing.Size(165, 29)
		Me.B_BoardOption.TabIndex = 13
		Me.B_BoardOption.Text = "Board Options"
		Me.B_BoardOption.UseVisualStyleBackColor = True
		'
		'B_PartUsage
		'
		Me.B_PartUsage.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
		Me.B_PartUsage.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!)
		Me.B_PartUsage.Location = New System.Drawing.Point(20, 165)
		Me.B_PartUsage.Name = "B_PartUsage"
		Me.B_PartUsage.Size = New System.Drawing.Size(165, 29)
		Me.B_PartUsage.TabIndex = 12
		Me.B_PartUsage.Text = "Part Usage"
		Me.B_PartUsage.UseVisualStyleBackColor = True
		'
		'ImportData_Button
		'
		Me.ImportData_Button.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
		Me.ImportData_Button.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!)
		Me.ImportData_Button.Location = New System.Drawing.Point(20, 95)
		Me.ImportData_Button.Name = "ImportData_Button"
		Me.ImportData_Button.Size = New System.Drawing.Size(165, 29)
		Me.ImportData_Button.TabIndex = 4
		Me.ImportData_Button.Text = "Import Data"
		Me.ImportData_Button.UseVisualStyleBackColor = True
		'
		'B_FindStockNumber
		'
		Me.B_FindStockNumber.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
		Me.B_FindStockNumber.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!)
		Me.B_FindStockNumber.Location = New System.Drawing.Point(20, 235)
		Me.B_FindStockNumber.Name = "B_FindStockNumber"
		Me.B_FindStockNumber.Size = New System.Drawing.Size(165, 29)
		Me.B_FindStockNumber.TabIndex = 15
		Me.B_FindStockNumber.Text = "Find Item Number"
		Me.B_FindStockNumber.UseVisualStyleBackColor = True
		'
		'B_MasterWindow
		'
		Me.B_MasterWindow.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
		Me.B_MasterWindow.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!)
		Me.B_MasterWindow.Location = New System.Drawing.Point(20, 594)
		Me.B_MasterWindow.Name = "B_MasterWindow"
		Me.B_MasterWindow.Size = New System.Drawing.Size(165, 29)
		Me.B_MasterWindow.TabIndex = 17
		Me.B_MasterWindow.Text = "Master Control"
		Me.B_MasterWindow.UseVisualStyleBackColor = True
		'
		'B_Settings
		'
		Me.B_Settings.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!)
		Me.B_Settings.Location = New System.Drawing.Point(20, 629)
		Me.B_Settings.Name = "B_Settings"
		Me.B_Settings.Size = New System.Drawing.Size(165, 30)
		Me.B_Settings.TabIndex = 18
		Me.B_Settings.Text = "Settings"
		Me.B_Settings.UseVisualStyleBackColor = True
		'
		'B_CompareRevisions
		'
		Me.B_CompareRevisions.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
		Me.B_CompareRevisions.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!)
		Me.B_CompareRevisions.Location = New System.Drawing.Point(20, 130)
		Me.B_CompareRevisions.Name = "B_CompareRevisions"
		Me.B_CompareRevisions.Size = New System.Drawing.Size(165, 29)
		Me.B_CompareRevisions.TabIndex = 11
		Me.B_CompareRevisions.Text = "Compare Revisions"
		Me.B_CompareRevisions.UseVisualStyleBackColor = True
		'
		'B_Close
		'
		Me.B_Close.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!)
		Me.B_Close.Location = New System.Drawing.Point(295, 14)
		Me.B_Close.Name = "B_Close"
		Me.B_Close.Size = New System.Drawing.Size(82, 29)
		Me.B_Close.TabIndex = 3
		Me.B_Close.Text = "Close"
		Me.B_Close.UseVisualStyleBackColor = True
		'
		'B_Cascade
		'
		Me.B_Cascade.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
		Me.B_Cascade.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!)
		Me.B_Cascade.Location = New System.Drawing.Point(6, 14)
		Me.B_Cascade.Name = "B_Cascade"
		Me.B_Cascade.Size = New System.Drawing.Size(107, 29)
		Me.B_Cascade.TabIndex = 0
		Me.B_Cascade.Text = "Cascade"
		Me.B_Cascade.UseVisualStyleBackColor = True
		'
		'B_Minimize
		'
		Me.B_Minimize.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!)
		Me.B_Minimize.Location = New System.Drawing.Point(207, 14)
		Me.B_Minimize.Name = "B_Minimize"
		Me.B_Minimize.Size = New System.Drawing.Size(82, 29)
		Me.B_Minimize.TabIndex = 2
		Me.B_Minimize.Text = "Minimize"
		Me.B_Minimize.UseVisualStyleBackColor = True
		'
		'B_Tile
		'
		Me.B_Tile.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!)
		Me.B_Tile.Location = New System.Drawing.Point(119, 14)
		Me.B_Tile.Name = "B_Tile"
		Me.B_Tile.Size = New System.Drawing.Size(82, 29)
		Me.B_Tile.TabIndex = 1
		Me.B_Tile.Text = "Tile"
		Me.B_Tile.UseVisualStyleBackColor = True
		'
		'B_Exit
		'
		Me.B_Exit.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
		Me.B_Exit.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!)
		Me.B_Exit.Location = New System.Drawing.Point(393, 13)
		Me.B_Exit.Name = "B_Exit"
		Me.B_Exit.Size = New System.Drawing.Size(108, 29)
		Me.B_Exit.TabIndex = 0
		Me.B_Exit.Text = "Exit Program"
		Me.B_Exit.UseVisualStyleBackColor = True
		'
		'GroupBox1
		'
		Me.GroupBox1.Controls.Add(Me.B_Close)
		Me.GroupBox1.Controls.Add(Me.B_Cascade)
		Me.GroupBox1.Controls.Add(Me.B_Minimize)
		Me.GroupBox1.Controls.Add(Me.B_Tile)
		Me.GroupBox1.Location = New System.Drawing.Point(4, -1)
		Me.GroupBox1.Name = "GroupBox1"
		Me.GroupBox1.Size = New System.Drawing.Size(383, 50)
		Me.GroupBox1.TabIndex = 20
		Me.GroupBox1.TabStop = False
		Me.GroupBox1.Text = "Window Buttons"
		'
		'Panel1
		'
		Me.Panel1.AutoScroll = True
		Me.Panel1.BackColor = System.Drawing.Color.Silver
		Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.Panel1.Controls.Add(Me.B_Exit)
		Me.Panel1.Controls.Add(Me.GroupBox1)
		Me.Panel1.Dock = System.Windows.Forms.DockStyle.Top
		Me.Panel1.Location = New System.Drawing.Point(208, 0)
		Me.Panel1.Name = "Panel1"
		Me.Panel1.Size = New System.Drawing.Size(1141, 60)
		Me.Panel1.TabIndex = 8
		'
		'MenuMain
		'
		Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.ClientSize = New System.Drawing.Size(1349, 673)
		Me.Controls.Add(Me.Panel1)
		Me.Controls.Add(Me.P_Menu)
		Me.Name = "MenuMain"
		Me.Text = "Engineering Tool"
		Me.P_Menu.ResumeLayout(False)
		Me.P_Menu.PerformLayout()
		Me.GroupBox1.ResumeLayout(False)
		Me.Panel1.ResumeLayout(False)
		Me.ResumeLayout(False)

	End Sub

	Friend WithEvents P_Menu As Panel
	Friend WithEvents Label2 As Label
	Friend WithEvents L_LastALPHAImport As Label
	Friend WithEvents Label1 As Label
	Friend WithEvents L_LastImport As Label
	Friend WithEvents B_BoardOption As Button
	Friend WithEvents B_PartUsage As Button
	Friend WithEvents ImportData_Button As Button
	Friend WithEvents B_FindStockNumber As Button
	Friend WithEvents B_MasterWindow As Button
	Friend WithEvents B_Settings As Button
	Friend WithEvents B_CompareRevisions As Button
	Friend WithEvents B_Close As Button
	Friend WithEvents B_Cascade As Button
	Friend WithEvents B_Minimize As Button
	Friend WithEvents B_Tile As Button
	Friend WithEvents B_Exit As Button
	Friend WithEvents GroupBox1 As GroupBox
	Friend WithEvents Panel1 As Panel
End Class
