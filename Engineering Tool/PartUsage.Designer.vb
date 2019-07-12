<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class PartUsage
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
		Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
		Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
		Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
		Dim DataGridViewCellStyle4 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
		Me.RefreshMags_Button = New System.Windows.Forms.Button()
		Me.MagInfo_CheckBox = New System.Windows.Forms.CheckBox()
		Me.Cost_CheckBox = New System.Windows.Forms.CheckBox()
		Me.AllVendors_CheckBox = New System.Windows.Forms.CheckBox()
		Me.Label3 = New System.Windows.Forms.Label()
		Me.Label2 = New System.Windows.Forms.Label()
		Me.Add_Button = New System.Windows.Forms.Button()
		Me.Clear_Button = New System.Windows.Forms.Button()
		Me.Remove_Button = New System.Windows.Forms.Button()
		Me.ChosenBoards_ListBox = New System.Windows.Forms.ListBox()
		Me.Boards_ListBox = New System.Windows.Forms.ListBox()
		Me.Label1 = New System.Windows.Forms.Label()
		Me.Excel_Button = New System.Windows.Forms.Button()
		Me.B_GenerateReport = New System.Windows.Forms.Button()
		Me.DGV_PartUsage = New System.Windows.Forms.DataGridView()
		Me.B_Close = New System.Windows.Forms.Button()
		CType(Me.DGV_PartUsage, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.SuspendLayout()
		'
		'RefreshMags_Button
		'
		Me.RefreshMags_Button.AutoSize = True
		Me.RefreshMags_Button.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
		Me.RefreshMags_Button.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!)
		Me.RefreshMags_Button.Location = New System.Drawing.Point(598, 13)
		Me.RefreshMags_Button.Name = "RefreshMags_Button"
		Me.RefreshMags_Button.Size = New System.Drawing.Size(119, 30)
		Me.RefreshMags_Button.TabIndex = 63
		Me.RefreshMags_Button.Text = "Refresh Mags"
		Me.RefreshMags_Button.UseVisualStyleBackColor = True
		'
		'MagInfo_CheckBox
		'
		Me.MagInfo_CheckBox.AutoSize = True
		Me.MagInfo_CheckBox.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.MagInfo_CheckBox.Location = New System.Drawing.Point(501, 17)
		Me.MagInfo_CheckBox.Name = "MagInfo_CheckBox"
		Me.MagInfo_CheckBox.Size = New System.Drawing.Size(91, 24)
		Me.MagInfo_CheckBox.TabIndex = 62
		Me.MagInfo_CheckBox.Text = "Mag Info"
		Me.MagInfo_CheckBox.UseVisualStyleBackColor = True
		'
		'Cost_CheckBox
		'
		Me.Cost_CheckBox.AutoSize = True
		Me.Cost_CheckBox.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Cost_CheckBox.Location = New System.Drawing.Point(434, 17)
		Me.Cost_CheckBox.Name = "Cost_CheckBox"
		Me.Cost_CheckBox.Size = New System.Drawing.Size(61, 24)
		Me.Cost_CheckBox.TabIndex = 61
		Me.Cost_CheckBox.Text = "Cost"
		Me.Cost_CheckBox.UseVisualStyleBackColor = True
		'
		'AllVendors_CheckBox
		'
		Me.AllVendors_CheckBox.AutoSize = True
		Me.AllVendors_CheckBox.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.AllVendors_CheckBox.Location = New System.Drawing.Point(319, 17)
		Me.AllVendors_CheckBox.Name = "AllVendors_CheckBox"
		Me.AllVendors_CheckBox.Size = New System.Drawing.Size(109, 24)
		Me.AllVendors_CheckBox.TabIndex = 60
		Me.AllVendors_CheckBox.Text = "All Vendors"
		Me.AllVendors_CheckBox.UseVisualStyleBackColor = True
		'
		'Label3
		'
		Me.Label3.AutoSize = True
		Me.Label3.Font = New System.Drawing.Font("Consolas", 12.0!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Underline), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label3.Location = New System.Drawing.Point(13, 10)
		Me.Label3.Name = "Label3"
		Me.Label3.Size = New System.Drawing.Size(99, 19)
		Me.Label3.TabIndex = 59
		Me.Label3.Text = "Part Usage"
		'
		'Label2
		'
		Me.Label2.AutoSize = True
		Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label2.Location = New System.Drawing.Point(13, 323)
		Me.Label2.Name = "Label2"
		Me.Label2.Size = New System.Drawing.Size(132, 20)
		Me.Label2.TabIndex = 58
		Me.Label2.Text = "Chosen Boards"
		'
		'Add_Button
		'
		Me.Add_Button.AutoSize = True
		Me.Add_Button.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
		Me.Add_Button.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!)
		Me.Add_Button.Location = New System.Drawing.Point(17, 290)
		Me.Add_Button.Name = "Add_Button"
		Me.Add_Button.Size = New System.Drawing.Size(48, 30)
		Me.Add_Button.TabIndex = 54
		Me.Add_Button.Text = "Add"
		Me.Add_Button.UseVisualStyleBackColor = True
		'
		'Clear_Button
		'
		Me.Clear_Button.AutoSize = True
		Me.Clear_Button.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
		Me.Clear_Button.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!)
		Me.Clear_Button.Location = New System.Drawing.Point(257, 556)
		Me.Clear_Button.Name = "Clear_Button"
		Me.Clear_Button.Size = New System.Drawing.Size(56, 30)
		Me.Clear_Button.TabIndex = 57
		Me.Clear_Button.Text = "Clear"
		Me.Clear_Button.UseVisualStyleBackColor = True
		'
		'Remove_Button
		'
		Me.Remove_Button.AutoSize = True
		Me.Remove_Button.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
		Me.Remove_Button.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!)
		Me.Remove_Button.Location = New System.Drawing.Point(17, 556)
		Me.Remove_Button.Name = "Remove_Button"
		Me.Remove_Button.Size = New System.Drawing.Size(78, 30)
		Me.Remove_Button.TabIndex = 56
		Me.Remove_Button.Text = "Remove"
		Me.Remove_Button.UseVisualStyleBackColor = True
		'
		'ChosenBoards_ListBox
		'
		Me.ChosenBoards_ListBox.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.ChosenBoards_ListBox.FormattingEnabled = True
		Me.ChosenBoards_ListBox.HorizontalScrollbar = True
		Me.ChosenBoards_ListBox.ItemHeight = 20
		Me.ChosenBoards_ListBox.Location = New System.Drawing.Point(17, 346)
		Me.ChosenBoards_ListBox.Name = "ChosenBoards_ListBox"
		Me.ChosenBoards_ListBox.ScrollAlwaysVisible = True
		Me.ChosenBoards_ListBox.Size = New System.Drawing.Size(296, 204)
		Me.ChosenBoards_ListBox.TabIndex = 55
		'
		'Boards_ListBox
		'
		Me.Boards_ListBox.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Boards_ListBox.FormattingEnabled = True
		Me.Boards_ListBox.HorizontalScrollbar = True
		Me.Boards_ListBox.ItemHeight = 20
		Me.Boards_ListBox.Location = New System.Drawing.Point(17, 80)
		Me.Boards_ListBox.Name = "Boards_ListBox"
		Me.Boards_ListBox.ScrollAlwaysVisible = True
		Me.Boards_ListBox.Size = New System.Drawing.Size(296, 204)
		Me.Boards_ListBox.TabIndex = 53
		'
		'Label1
		'
		Me.Label1.AutoSize = True
		Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label1.Location = New System.Drawing.Point(13, 49)
		Me.Label1.Name = "Label1"
		Me.Label1.Size = New System.Drawing.Size(194, 20)
		Me.Label1.TabIndex = 52
		Me.Label1.Text = "Boards to Choose from"
		'
		'Excel_Button
		'
		Me.Excel_Button.AutoSize = True
		Me.Excel_Button.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
		Me.Excel_Button.Enabled = False
		Me.Excel_Button.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Excel_Button.Location = New System.Drawing.Point(869, 13)
		Me.Excel_Button.Name = "Excel_Button"
		Me.Excel_Button.Size = New System.Drawing.Size(109, 30)
		Me.Excel_Button.TabIndex = 49
		Me.Excel_Button.Text = "Create Excel"
		Me.Excel_Button.UseVisualStyleBackColor = True
		'
		'B_GenerateReport
		'
		Me.B_GenerateReport.AutoSize = True
		Me.B_GenerateReport.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
		Me.B_GenerateReport.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!)
		Me.B_GenerateReport.Location = New System.Drawing.Point(723, 13)
		Me.B_GenerateReport.Name = "B_GenerateReport"
		Me.B_GenerateReport.Size = New System.Drawing.Size(140, 30)
		Me.B_GenerateReport.TabIndex = 48
		Me.B_GenerateReport.Text = "Generate Report"
		Me.B_GenerateReport.UseVisualStyleBackColor = True
		'
		'DGV_PartUsage
		'
		Me.DGV_PartUsage.AllowUserToAddRows = False
		Me.DGV_PartUsage.AllowUserToDeleteRows = False
		Me.DGV_PartUsage.AllowUserToOrderColumns = True
		DataGridViewCellStyle1.BackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
		Me.DGV_PartUsage.AlternatingRowsDefaultCellStyle = DataGridViewCellStyle1
		Me.DGV_PartUsage.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
			Or System.Windows.Forms.AnchorStyles.Left) _
			Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
		Me.DGV_PartUsage.BackgroundColor = System.Drawing.SystemColors.Control
		DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
		DataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Control
		DataGridViewCellStyle2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		DataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.WindowText
		DataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight
		DataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText
		DataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
		Me.DGV_PartUsage.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle2
		Me.DGV_PartUsage.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
		DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
		DataGridViewCellStyle3.BackColor = System.Drawing.SystemColors.Window
		DataGridViewCellStyle3.Font = New System.Drawing.Font("Consolas", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		DataGridViewCellStyle3.ForeColor = System.Drawing.SystemColors.ControlText
		DataGridViewCellStyle3.SelectionBackColor = System.Drawing.SystemColors.Highlight
		DataGridViewCellStyle3.SelectionForeColor = System.Drawing.SystemColors.HighlightText
		DataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
		Me.DGV_PartUsage.DefaultCellStyle = DataGridViewCellStyle3
		Me.DGV_PartUsage.Location = New System.Drawing.Point(319, 49)
		Me.DGV_PartUsage.Name = "DGV_PartUsage"
		DataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
		DataGridViewCellStyle4.BackColor = System.Drawing.SystemColors.Control
		DataGridViewCellStyle4.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		DataGridViewCellStyle4.ForeColor = System.Drawing.SystemColors.WindowText
		DataGridViewCellStyle4.SelectionBackColor = System.Drawing.SystemColors.Highlight
		DataGridViewCellStyle4.SelectionForeColor = System.Drawing.SystemColors.HighlightText
		DataGridViewCellStyle4.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
		Me.DGV_PartUsage.RowHeadersDefaultCellStyle = DataGridViewCellStyle4
		Me.DGV_PartUsage.Size = New System.Drawing.Size(722, 536)
		Me.DGV_PartUsage.TabIndex = 51
		'
		'B_Close
		'
		Me.B_Close.AutoSize = True
		Me.B_Close.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
		Me.B_Close.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!)
		Me.B_Close.Location = New System.Drawing.Point(984, 13)
		Me.B_Close.Name = "B_Close"
		Me.B_Close.Size = New System.Drawing.Size(59, 30)
		Me.B_Close.TabIndex = 50
		Me.B_Close.Text = "Close"
		Me.B_Close.UseVisualStyleBackColor = True
		'
		'PartUsage
		'
		Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.ClientSize = New System.Drawing.Size(1056, 596)
		Me.Controls.Add(Me.RefreshMags_Button)
		Me.Controls.Add(Me.MagInfo_CheckBox)
		Me.Controls.Add(Me.Cost_CheckBox)
		Me.Controls.Add(Me.AllVendors_CheckBox)
		Me.Controls.Add(Me.Label3)
		Me.Controls.Add(Me.Label2)
		Me.Controls.Add(Me.Add_Button)
		Me.Controls.Add(Me.Clear_Button)
		Me.Controls.Add(Me.Remove_Button)
		Me.Controls.Add(Me.ChosenBoards_ListBox)
		Me.Controls.Add(Me.Boards_ListBox)
		Me.Controls.Add(Me.Label1)
		Me.Controls.Add(Me.Excel_Button)
		Me.Controls.Add(Me.B_GenerateReport)
		Me.Controls.Add(Me.DGV_PartUsage)
		Me.Controls.Add(Me.B_Close)
		Me.Name = "PartUsage"
		Me.Text = "Part Usage"
		CType(Me.DGV_PartUsage, System.ComponentModel.ISupportInitialize).EndInit()
		Me.ResumeLayout(False)
		Me.PerformLayout()

	End Sub

	Friend WithEvents RefreshMags_Button As Button
	Friend WithEvents MagInfo_CheckBox As CheckBox
	Friend WithEvents Cost_CheckBox As CheckBox
	Friend WithEvents AllVendors_CheckBox As CheckBox
	Friend WithEvents Label3 As Label
	Friend WithEvents Label2 As Label
	Friend WithEvents Add_Button As Button
	Friend WithEvents Clear_Button As Button
	Friend WithEvents Remove_Button As Button
	Friend WithEvents ChosenBoards_ListBox As ListBox
	Friend WithEvents Boards_ListBox As ListBox
	Friend WithEvents Label1 As Label
	Friend WithEvents Excel_Button As Button
	Friend WithEvents B_GenerateReport As Button
	Friend WithEvents DGV_PartUsage As DataGridView
	Friend WithEvents B_Close As Button
End Class
