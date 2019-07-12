<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class BoardOptions
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
		Me.Label3 = New System.Windows.Forms.Label()
		Me.Label2 = New System.Windows.Forms.Label()
		Me.Add_Button = New System.Windows.Forms.Button()
		Me.Clear_Button = New System.Windows.Forms.Button()
		Me.Remove_Button = New System.Windows.Forms.Button()
		Me.ChosenBoards_ListBox = New System.Windows.Forms.ListBox()
		Me.Boards_ListBox = New System.Windows.Forms.ListBox()
		Me.Excel_Button = New System.Windows.Forms.Button()
		Me.GenerateReport_Button = New System.Windows.Forms.Button()
		Me.BoardOptions_DGV = New System.Windows.Forms.DataGridView()
		Me.Label1 = New System.Windows.Forms.Label()
		Me.Close_Button = New System.Windows.Forms.Button()
		CType(Me.BoardOptions_DGV, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.SuspendLayout()
		'
		'Label3
		'
		Me.Label3.AutoSize = True
		Me.Label3.Font = New System.Drawing.Font("Consolas", 12.0!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Underline), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label3.Location = New System.Drawing.Point(15, 11)
		Me.Label3.Name = "Label3"
		Me.Label3.Size = New System.Drawing.Size(126, 19)
		Me.Label3.TabIndex = 48
		Me.Label3.Text = "Board Options"
		'
		'Label2
		'
		Me.Label2.AutoSize = True
		Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label2.Location = New System.Drawing.Point(15, 316)
		Me.Label2.Name = "Label2"
		Me.Label2.Size = New System.Drawing.Size(132, 20)
		Me.Label2.TabIndex = 47
		Me.Label2.Text = "Chosen Boards"
		'
		'Add_Button
		'
		Me.Add_Button.AutoSize = True
		Me.Add_Button.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
		Me.Add_Button.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!)
		Me.Add_Button.Location = New System.Drawing.Point(19, 283)
		Me.Add_Button.Name = "Add_Button"
		Me.Add_Button.Size = New System.Drawing.Size(48, 30)
		Me.Add_Button.TabIndex = 43
		Me.Add_Button.Text = "Add"
		Me.Add_Button.UseVisualStyleBackColor = True
		'
		'Clear_Button
		'
		Me.Clear_Button.AutoSize = True
		Me.Clear_Button.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
		Me.Clear_Button.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!)
		Me.Clear_Button.Location = New System.Drawing.Point(259, 549)
		Me.Clear_Button.Name = "Clear_Button"
		Me.Clear_Button.Size = New System.Drawing.Size(56, 30)
		Me.Clear_Button.TabIndex = 46
		Me.Clear_Button.Text = "Clear"
		Me.Clear_Button.UseVisualStyleBackColor = True
		'
		'Remove_Button
		'
		Me.Remove_Button.AutoSize = True
		Me.Remove_Button.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
		Me.Remove_Button.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!)
		Me.Remove_Button.Location = New System.Drawing.Point(19, 549)
		Me.Remove_Button.Name = "Remove_Button"
		Me.Remove_Button.Size = New System.Drawing.Size(78, 30)
		Me.Remove_Button.TabIndex = 45
		Me.Remove_Button.Text = "Remove"
		Me.Remove_Button.UseVisualStyleBackColor = True
		'
		'ChosenBoards_ListBox
		'
		Me.ChosenBoards_ListBox.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.ChosenBoards_ListBox.FormattingEnabled = True
		Me.ChosenBoards_ListBox.HorizontalScrollbar = True
		Me.ChosenBoards_ListBox.ItemHeight = 20
		Me.ChosenBoards_ListBox.Location = New System.Drawing.Point(19, 339)
		Me.ChosenBoards_ListBox.Name = "ChosenBoards_ListBox"
		Me.ChosenBoards_ListBox.ScrollAlwaysVisible = True
		Me.ChosenBoards_ListBox.Size = New System.Drawing.Size(296, 204)
		Me.ChosenBoards_ListBox.TabIndex = 44
		'
		'Boards_ListBox
		'
		Me.Boards_ListBox.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Boards_ListBox.FormattingEnabled = True
		Me.Boards_ListBox.HorizontalScrollbar = True
		Me.Boards_ListBox.ItemHeight = 20
		Me.Boards_ListBox.Location = New System.Drawing.Point(19, 73)
		Me.Boards_ListBox.Name = "Boards_ListBox"
		Me.Boards_ListBox.ScrollAlwaysVisible = True
		Me.Boards_ListBox.Size = New System.Drawing.Size(296, 204)
		Me.Boards_ListBox.TabIndex = 42
		'
		'Excel_Button
		'
		Me.Excel_Button.AutoSize = True
		Me.Excel_Button.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
		Me.Excel_Button.Enabled = False
		Me.Excel_Button.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Excel_Button.Location = New System.Drawing.Point(467, 14)
		Me.Excel_Button.Name = "Excel_Button"
		Me.Excel_Button.Size = New System.Drawing.Size(109, 30)
		Me.Excel_Button.TabIndex = 38
		Me.Excel_Button.Text = "Create Excel"
		Me.Excel_Button.UseVisualStyleBackColor = True
		'
		'GenerateReport_Button
		'
		Me.GenerateReport_Button.AutoSize = True
		Me.GenerateReport_Button.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
		Me.GenerateReport_Button.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!)
		Me.GenerateReport_Button.Location = New System.Drawing.Point(321, 14)
		Me.GenerateReport_Button.Name = "GenerateReport_Button"
		Me.GenerateReport_Button.Size = New System.Drawing.Size(140, 30)
		Me.GenerateReport_Button.TabIndex = 37
		Me.GenerateReport_Button.Text = "Generate Report"
		Me.GenerateReport_Button.UseVisualStyleBackColor = True
		'
		'BoardOptions_DGV
		'
		Me.BoardOptions_DGV.AllowUserToAddRows = False
		Me.BoardOptions_DGV.AllowUserToDeleteRows = False
		Me.BoardOptions_DGV.AllowUserToOrderColumns = True
		DataGridViewCellStyle1.BackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
		Me.BoardOptions_DGV.AlternatingRowsDefaultCellStyle = DataGridViewCellStyle1
		Me.BoardOptions_DGV.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
			Or System.Windows.Forms.AnchorStyles.Left) _
			Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
		Me.BoardOptions_DGV.BackgroundColor = System.Drawing.SystemColors.Control
		DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
		DataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Control
		DataGridViewCellStyle2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		DataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.WindowText
		DataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight
		DataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText
		DataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
		Me.BoardOptions_DGV.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle2
		Me.BoardOptions_DGV.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
		DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
		DataGridViewCellStyle3.BackColor = System.Drawing.SystemColors.Window
		DataGridViewCellStyle3.Font = New System.Drawing.Font("Consolas", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		DataGridViewCellStyle3.ForeColor = System.Drawing.SystemColors.ControlText
		DataGridViewCellStyle3.SelectionBackColor = System.Drawing.SystemColors.Highlight
		DataGridViewCellStyle3.SelectionForeColor = System.Drawing.SystemColors.HighlightText
		DataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
		Me.BoardOptions_DGV.DefaultCellStyle = DataGridViewCellStyle3
		Me.BoardOptions_DGV.Location = New System.Drawing.Point(321, 50)
		Me.BoardOptions_DGV.Name = "BoardOptions_DGV"
		DataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
		DataGridViewCellStyle4.BackColor = System.Drawing.SystemColors.Control
		DataGridViewCellStyle4.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		DataGridViewCellStyle4.ForeColor = System.Drawing.SystemColors.WindowText
		DataGridViewCellStyle4.SelectionBackColor = System.Drawing.SystemColors.Highlight
		DataGridViewCellStyle4.SelectionForeColor = System.Drawing.SystemColors.HighlightText
		DataGridViewCellStyle4.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
		Me.BoardOptions_DGV.RowHeadersDefaultCellStyle = DataGridViewCellStyle4
		Me.BoardOptions_DGV.Size = New System.Drawing.Size(627, 530)
		Me.BoardOptions_DGV.TabIndex = 40
		'
		'Label1
		'
		Me.Label1.AutoSize = True
		Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label1.Location = New System.Drawing.Point(15, 50)
		Me.Label1.Name = "Label1"
		Me.Label1.Size = New System.Drawing.Size(194, 20)
		Me.Label1.TabIndex = 41
		Me.Label1.Text = "Boards to Choose from"
		'
		'Close_Button
		'
		Me.Close_Button.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
		Me.Close_Button.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!)
		Me.Close_Button.Location = New System.Drawing.Point(582, 14)
		Me.Close_Button.Name = "Close_Button"
		Me.Close_Button.Size = New System.Drawing.Size(78, 30)
		Me.Close_Button.TabIndex = 39
		Me.Close_Button.Text = "Close"
		Me.Close_Button.UseVisualStyleBackColor = True
		'
		'BoardOptions
		'
		Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.ClientSize = New System.Drawing.Size(963, 590)
		Me.Controls.Add(Me.Label3)
		Me.Controls.Add(Me.Label2)
		Me.Controls.Add(Me.Add_Button)
		Me.Controls.Add(Me.Clear_Button)
		Me.Controls.Add(Me.Remove_Button)
		Me.Controls.Add(Me.ChosenBoards_ListBox)
		Me.Controls.Add(Me.Boards_ListBox)
		Me.Controls.Add(Me.Excel_Button)
		Me.Controls.Add(Me.GenerateReport_Button)
		Me.Controls.Add(Me.BoardOptions_DGV)
		Me.Controls.Add(Me.Label1)
		Me.Controls.Add(Me.Close_Button)
		Me.Name = "BoardOptions"
		Me.Text = "Board Options"
		CType(Me.BoardOptions_DGV, System.ComponentModel.ISupportInitialize).EndInit()
		Me.ResumeLayout(False)
		Me.PerformLayout()

	End Sub

	Friend WithEvents Label3 As Label
	Friend WithEvents Label2 As Label
	Friend WithEvents Add_Button As Button
	Friend WithEvents Clear_Button As Button
	Friend WithEvents Remove_Button As Button
	Friend WithEvents ChosenBoards_ListBox As ListBox
	Friend WithEvents Boards_ListBox As ListBox
	Friend WithEvents Excel_Button As Button
	Friend WithEvents GenerateReport_Button As Button
	Friend WithEvents BoardOptions_DGV As DataGridView
	Friend WithEvents Label1 As Label
	Friend WithEvents Close_Button As Button
End Class
