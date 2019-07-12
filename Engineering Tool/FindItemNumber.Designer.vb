<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FindItemNumber
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
		Dim DataGridViewCellStyle5 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
		Dim DataGridViewCellStyle6 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
		Dim DataGridViewCellStyle7 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
		Dim DataGridViewCellStyle8 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
		Me.Label3 = New System.Windows.Forms.Label()
		Me.TB_Search2 = New System.Windows.Forms.TextBox()
		Me.CB_Search2 = New System.Windows.Forms.ComboBox()
		Me.TB_Search = New System.Windows.Forms.TextBox()
		Me.B_Search = New System.Windows.Forms.Button()
		Me.CB_Search = New System.Windows.Forms.ComboBox()
		Me.DGV_Results = New System.Windows.Forms.DataGridView()
		Me.L_Found = New System.Windows.Forms.Label()
		Me.B_Close = New System.Windows.Forms.Button()
		CType(Me.DGV_Results, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.SuspendLayout()
		'
		'Label3
		'
		Me.Label3.AutoSize = True
		Me.Label3.Font = New System.Drawing.Font("Consolas", 12.0!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Underline), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label3.Location = New System.Drawing.Point(12, 10)
		Me.Label3.Name = "Label3"
		Me.Label3.Size = New System.Drawing.Size(153, 19)
		Me.Label3.TabIndex = 46
		Me.Label3.Text = "Find Item Number"
		'
		'TB_Search2
		'
		Me.TB_Search2.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.TB_Search2.Location = New System.Drawing.Point(230, 66)
		Me.TB_Search2.Name = "TB_Search2"
		Me.TB_Search2.Size = New System.Drawing.Size(254, 26)
		Me.TB_Search2.TabIndex = 41
		'
		'CB_Search2
		'
		Me.CB_Search2.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
		Me.CB_Search2.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.CB_Search2.FormattingEnabled = True
		Me.CB_Search2.Location = New System.Drawing.Point(12, 66)
		Me.CB_Search2.Name = "CB_Search2"
		Me.CB_Search2.Size = New System.Drawing.Size(212, 28)
		Me.CB_Search2.TabIndex = 40
		'
		'TB_Search
		'
		Me.TB_Search.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.TB_Search.Location = New System.Drawing.Point(230, 32)
		Me.TB_Search.Name = "TB_Search"
		Me.TB_Search.Size = New System.Drawing.Size(254, 26)
		Me.TB_Search.TabIndex = 39
		'
		'B_Search
		'
		Me.B_Search.AutoSize = True
		Me.B_Search.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
		Me.B_Search.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.B_Search.Location = New System.Drawing.Point(490, 30)
		Me.B_Search.Name = "B_Search"
		Me.B_Search.Size = New System.Drawing.Size(70, 30)
		Me.B_Search.TabIndex = 42
		Me.B_Search.Text = "Search"
		Me.B_Search.UseVisualStyleBackColor = True
		'
		'CB_Search
		'
		Me.CB_Search.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
		Me.CB_Search.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.CB_Search.FormattingEnabled = True
		Me.CB_Search.Location = New System.Drawing.Point(12, 32)
		Me.CB_Search.Name = "CB_Search"
		Me.CB_Search.Size = New System.Drawing.Size(212, 28)
		Me.CB_Search.TabIndex = 38
		'
		'DGV_Results
		'
		Me.DGV_Results.AllowUserToAddRows = False
		Me.DGV_Results.AllowUserToDeleteRows = False
		Me.DGV_Results.AllowUserToOrderColumns = True
		DataGridViewCellStyle5.BackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
		Me.DGV_Results.AlternatingRowsDefaultCellStyle = DataGridViewCellStyle5
		Me.DGV_Results.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
			Or System.Windows.Forms.AnchorStyles.Left) _
			Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
		Me.DGV_Results.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
		Me.DGV_Results.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
		Me.DGV_Results.BackgroundColor = System.Drawing.SystemColors.Control
		DataGridViewCellStyle6.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
		DataGridViewCellStyle6.BackColor = System.Drawing.SystemColors.Control
		DataGridViewCellStyle6.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		DataGridViewCellStyle6.ForeColor = System.Drawing.SystemColors.WindowText
		DataGridViewCellStyle6.SelectionBackColor = System.Drawing.SystemColors.Highlight
		DataGridViewCellStyle6.SelectionForeColor = System.Drawing.SystemColors.HighlightText
		DataGridViewCellStyle6.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
		Me.DGV_Results.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle6
		Me.DGV_Results.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
		DataGridViewCellStyle7.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
		DataGridViewCellStyle7.BackColor = System.Drawing.SystemColors.Window
		DataGridViewCellStyle7.Font = New System.Drawing.Font("Consolas", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		DataGridViewCellStyle7.ForeColor = System.Drawing.SystemColors.ControlText
		DataGridViewCellStyle7.SelectionBackColor = System.Drawing.SystemColors.Highlight
		DataGridViewCellStyle7.SelectionForeColor = System.Drawing.SystemColors.HighlightText
		DataGridViewCellStyle7.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
		Me.DGV_Results.DefaultCellStyle = DataGridViewCellStyle7
		Me.DGV_Results.Location = New System.Drawing.Point(12, 100)
		Me.DGV_Results.Name = "DGV_Results"
		DataGridViewCellStyle8.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
		DataGridViewCellStyle8.BackColor = System.Drawing.SystemColors.Control
		DataGridViewCellStyle8.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		DataGridViewCellStyle8.ForeColor = System.Drawing.SystemColors.WindowText
		DataGridViewCellStyle8.SelectionBackColor = System.Drawing.SystemColors.Highlight
		DataGridViewCellStyle8.SelectionForeColor = System.Drawing.SystemColors.HighlightText
		DataGridViewCellStyle8.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
		Me.DGV_Results.RowHeadersDefaultCellStyle = DataGridViewCellStyle8
		Me.DGV_Results.Size = New System.Drawing.Size(766, 366)
		Me.DGV_Results.TabIndex = 45
		'
		'L_Found
		'
		Me.L_Found.AutoSize = True
		Me.L_Found.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.L_Found.Location = New System.Drawing.Point(486, 69)
		Me.L_Found.Name = "L_Found"
		Me.L_Found.Size = New System.Drawing.Size(259, 20)
		Me.L_Found.TabIndex = 44
		Me.L_Found.Text = "Found in _ Board(s). Results: _"
		'
		'B_Close
		'
		Me.B_Close.AutoSize = True
		Me.B_Close.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
		Me.B_Close.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!)
		Me.B_Close.Location = New System.Drawing.Point(719, 30)
		Me.B_Close.Name = "B_Close"
		Me.B_Close.Size = New System.Drawing.Size(59, 30)
		Me.B_Close.TabIndex = 43
		Me.B_Close.Text = "Close"
		Me.B_Close.UseVisualStyleBackColor = True
		'
		'FindItemNumber
		'
		Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.ClientSize = New System.Drawing.Size(790, 477)
		Me.Controls.Add(Me.Label3)
		Me.Controls.Add(Me.TB_Search2)
		Me.Controls.Add(Me.CB_Search2)
		Me.Controls.Add(Me.TB_Search)
		Me.Controls.Add(Me.B_Search)
		Me.Controls.Add(Me.CB_Search)
		Me.Controls.Add(Me.DGV_Results)
		Me.Controls.Add(Me.L_Found)
		Me.Controls.Add(Me.B_Close)
		Me.Name = "FindItemNumber"
		Me.Text = "Find Item Number"
		CType(Me.DGV_Results, System.ComponentModel.ISupportInitialize).EndInit()
		Me.ResumeLayout(False)
		Me.PerformLayout()

	End Sub

	Friend WithEvents Label3 As Label
	Friend WithEvents TB_Search2 As TextBox
	Friend WithEvents CB_Search2 As ComboBox
	Friend WithEvents TB_Search As TextBox
	Friend WithEvents B_Search As Button
	Friend WithEvents CB_Search As ComboBox
	Friend WithEvents DGV_Results As DataGridView
	Friend WithEvents L_Found As Label
	Friend WithEvents B_Close As Button
End Class
