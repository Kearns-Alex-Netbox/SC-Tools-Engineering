<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class CompareRevisions
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
		Dim DataGridViewCellStyle5 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
		Dim DataGridViewCellStyle6 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
		Dim DataGridViewCellStyle7 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
		Dim DataGridViewCellStyle8 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
		Me.Label3 = New System.Windows.Forms.Label()
		Me.DGV_Board2_Quantity = New System.Windows.Forms.DataGridView()
		Me.DGV_Board1_Quantity = New System.Windows.Forms.DataGridView()
		Me.L_Board1Quantities = New System.Windows.Forms.Label()
		Me.L_Board2Quantities = New System.Windows.Forms.Label()
		Me.Label8 = New System.Windows.Forms.Label()
		Me.Label4 = New System.Windows.Forms.Label()
		Me.Label9 = New System.Windows.Forms.Label()
		Me.CB_Board2 = New System.Windows.Forms.ComboBox()
		Me.L_Title = New System.Windows.Forms.Label()
		Me.Label2 = New System.Windows.Forms.Label()
		Me.CB_Source = New System.Windows.Forms.ComboBox()
		Me.Excel_Button = New System.Windows.Forms.Button()
		Me.GenerateReport_Button = New System.Windows.Forms.Button()
		Me.Label1 = New System.Windows.Forms.Label()
		Me.CB_Board1 = New System.Windows.Forms.ComboBox()
		Me.Close_Button = New System.Windows.Forms.Button()
		CType(Me.DGV_Board2_Quantity, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.DGV_Board1_Quantity, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.SuspendLayout()
		'
		'Label3
		'
		Me.Label3.AutoSize = True
		Me.Label3.Font = New System.Drawing.Font("Consolas", 12.0!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Underline), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label3.Location = New System.Drawing.Point(8, 13)
		Me.Label3.Name = "Label3"
		Me.Label3.Size = New System.Drawing.Size(162, 19)
		Me.Label3.TabIndex = 54
		Me.Label3.Text = "Compare Revisions"
		'
		'DGV_Board2_Quantity
		'
		Me.DGV_Board2_Quantity.AllowUserToAddRows = False
		Me.DGV_Board2_Quantity.AllowUserToDeleteRows = False
		DataGridViewCellStyle1.BackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
		Me.DGV_Board2_Quantity.AlternatingRowsDefaultCellStyle = DataGridViewCellStyle1
		Me.DGV_Board2_Quantity.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
			Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
		Me.DGV_Board2_Quantity.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
		Me.DGV_Board2_Quantity.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
		Me.DGV_Board2_Quantity.BackgroundColor = System.Drawing.SystemColors.Control
		Me.DGV_Board2_Quantity.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
		DataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Control
		DataGridViewCellStyle2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		DataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.WindowText
		DataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight
		DataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText
		DataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
		Me.DGV_Board2_Quantity.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle2
		Me.DGV_Board2_Quantity.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
		DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
		DataGridViewCellStyle3.BackColor = System.Drawing.SystemColors.Window
		DataGridViewCellStyle3.Font = New System.Drawing.Font("Consolas", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		DataGridViewCellStyle3.ForeColor = System.Drawing.SystemColors.ControlText
		DataGridViewCellStyle3.SelectionBackColor = System.Drawing.SystemColors.Highlight
		DataGridViewCellStyle3.SelectionForeColor = System.Drawing.SystemColors.HighlightText
		DataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
		Me.DGV_Board2_Quantity.DefaultCellStyle = DataGridViewCellStyle3
		Me.DGV_Board2_Quantity.EditMode = System.Windows.Forms.DataGridViewEditMode.EditOnF2
		Me.DGV_Board2_Quantity.Location = New System.Drawing.Point(607, 123)
		Me.DGV_Board2_Quantity.Name = "DGV_Board2_Quantity"
		DataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
		DataGridViewCellStyle4.BackColor = System.Drawing.SystemColors.Control
		DataGridViewCellStyle4.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		DataGridViewCellStyle4.ForeColor = System.Drawing.SystemColors.WindowText
		DataGridViewCellStyle4.SelectionBackColor = System.Drawing.SystemColors.Highlight
		DataGridViewCellStyle4.SelectionForeColor = System.Drawing.SystemColors.HighlightText
		DataGridViewCellStyle4.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
		Me.DGV_Board2_Quantity.RowHeadersDefaultCellStyle = DataGridViewCellStyle4
		Me.DGV_Board2_Quantity.Size = New System.Drawing.Size(476, 445)
		Me.DGV_Board2_Quantity.TabIndex = 51
		'
		'DGV_Board1_Quantity
		'
		Me.DGV_Board1_Quantity.AllowUserToAddRows = False
		Me.DGV_Board1_Quantity.AllowUserToDeleteRows = False
		DataGridViewCellStyle5.BackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
		Me.DGV_Board1_Quantity.AlternatingRowsDefaultCellStyle = DataGridViewCellStyle5
		Me.DGV_Board1_Quantity.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
			Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
		Me.DGV_Board1_Quantity.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
		Me.DGV_Board1_Quantity.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
		Me.DGV_Board1_Quantity.BackgroundColor = System.Drawing.SystemColors.Control
		Me.DGV_Board1_Quantity.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		DataGridViewCellStyle6.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
		DataGridViewCellStyle6.BackColor = System.Drawing.SystemColors.Control
		DataGridViewCellStyle6.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		DataGridViewCellStyle6.ForeColor = System.Drawing.SystemColors.WindowText
		DataGridViewCellStyle6.SelectionBackColor = System.Drawing.SystemColors.Highlight
		DataGridViewCellStyle6.SelectionForeColor = System.Drawing.SystemColors.HighlightText
		DataGridViewCellStyle6.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
		Me.DGV_Board1_Quantity.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle6
		Me.DGV_Board1_Quantity.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
		DataGridViewCellStyle7.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
		DataGridViewCellStyle7.BackColor = System.Drawing.SystemColors.Window
		DataGridViewCellStyle7.Font = New System.Drawing.Font("Consolas", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		DataGridViewCellStyle7.ForeColor = System.Drawing.SystemColors.ControlText
		DataGridViewCellStyle7.SelectionBackColor = System.Drawing.SystemColors.Highlight
		DataGridViewCellStyle7.SelectionForeColor = System.Drawing.SystemColors.HighlightText
		DataGridViewCellStyle7.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
		Me.DGV_Board1_Quantity.DefaultCellStyle = DataGridViewCellStyle7
		Me.DGV_Board1_Quantity.EditMode = System.Windows.Forms.DataGridViewEditMode.EditOnF2
		Me.DGV_Board1_Quantity.Location = New System.Drawing.Point(12, 123)
		Me.DGV_Board1_Quantity.Name = "DGV_Board1_Quantity"
		DataGridViewCellStyle8.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
		DataGridViewCellStyle8.BackColor = System.Drawing.SystemColors.Control
		DataGridViewCellStyle8.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		DataGridViewCellStyle8.ForeColor = System.Drawing.SystemColors.WindowText
		DataGridViewCellStyle8.SelectionBackColor = System.Drawing.SystemColors.Highlight
		DataGridViewCellStyle8.SelectionForeColor = System.Drawing.SystemColors.HighlightText
		DataGridViewCellStyle8.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
		Me.DGV_Board1_Quantity.RowHeadersDefaultCellStyle = DataGridViewCellStyle8
		Me.DGV_Board1_Quantity.Size = New System.Drawing.Size(476, 445)
		Me.DGV_Board1_Quantity.TabIndex = 50
		'
		'L_Board1Quantities
		'
		Me.L_Board1Quantities.AutoSize = True
		Me.L_Board1Quantities.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.L_Board1Quantities.Location = New System.Drawing.Point(8, 100)
		Me.L_Board1Quantities.Name = "L_Board1Quantities"
		Me.L_Board1Quantities.Size = New System.Drawing.Size(167, 20)
		Me.L_Board1Quantities.TabIndex = 52
		Me.L_Board1Quantities.Text = "'Board 1' Quantities"
		'
		'L_Board2Quantities
		'
		Me.L_Board2Quantities.AutoSize = True
		Me.L_Board2Quantities.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.L_Board2Quantities.Location = New System.Drawing.Point(603, 100)
		Me.L_Board2Quantities.Name = "L_Board2Quantities"
		Me.L_Board2Quantities.Size = New System.Drawing.Size(167, 20)
		Me.L_Board2Quantities.TabIndex = 53
		Me.L_Board2Quantities.Text = "'Board 2' Quantities"
		'
		'Label8
		'
		Me.Label8.AutoSize = True
		Me.Label8.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label8.Location = New System.Drawing.Point(289, 59)
		Me.Label8.Name = "Label8"
		Me.Label8.Size = New System.Drawing.Size(47, 16)
		Me.Label8.TabIndex = 48
		Me.Label8.Text = "(older)"
		'
		'Label4
		'
		Me.Label4.AutoSize = True
		Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label4.Location = New System.Drawing.Point(573, 59)
		Me.Label4.Name = "Label4"
		Me.Label4.Size = New System.Drawing.Size(52, 16)
		Me.Label4.TabIndex = 47
		Me.Label4.Text = "(newer)"
		'
		'Label9
		'
		Me.Label9.AutoSize = True
		Me.Label9.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label9.Location = New System.Drawing.Point(573, 39)
		Me.Label9.Name = "Label9"
		Me.Label9.Size = New System.Drawing.Size(65, 20)
		Me.Label9.TabIndex = 42
		Me.Label9.Text = "Board 2"
		'
		'CB_Board2
		'
		Me.CB_Board2.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
		Me.CB_Board2.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.CB_Board2.FormattingEnabled = True
		Me.CB_Board2.Location = New System.Drawing.Point(644, 35)
		Me.CB_Board2.Name = "CB_Board2"
		Me.CB_Board2.Size = New System.Drawing.Size(208, 28)
		Me.CB_Board2.TabIndex = 43
		'
		'L_Title
		'
		Me.L_Title.AutoSize = True
		Me.L_Title.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.75!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Underline), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.L_Title.Location = New System.Drawing.Point(443, 75)
		Me.L_Title.Name = "L_Title"
		Me.L_Title.Size = New System.Drawing.Size(272, 25)
		Me.L_Title.TabIndex = 49
		Me.L_Title.Text = "'Board 1'   <->   'Board 2'"
		'
		'Label2
		'
		Me.Label2.AutoSize = True
		Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label2.Location = New System.Drawing.Point(8, 39)
		Me.Label2.Name = "Label2"
		Me.Label2.Size = New System.Drawing.Size(60, 20)
		Me.Label2.TabIndex = 38
		Me.Label2.Text = "Source"
		'
		'CB_Source
		'
		Me.CB_Source.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
		Me.CB_Source.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.CB_Source.FormattingEnabled = True
		Me.CB_Source.Location = New System.Drawing.Point(74, 35)
		Me.CB_Source.Name = "CB_Source"
		Me.CB_Source.Size = New System.Drawing.Size(208, 28)
		Me.CB_Source.TabIndex = 39
		'
		'Excel_Button
		'
		Me.Excel_Button.AutoSize = True
		Me.Excel_Button.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
		Me.Excel_Button.Enabled = False
		Me.Excel_Button.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!)
		Me.Excel_Button.Location = New System.Drawing.Point(1004, 34)
		Me.Excel_Button.Name = "Excel_Button"
		Me.Excel_Button.Size = New System.Drawing.Size(109, 30)
		Me.Excel_Button.TabIndex = 45
		Me.Excel_Button.Text = "Create Excel"
		Me.Excel_Button.UseVisualStyleBackColor = True
		'
		'GenerateReport_Button
		'
		Me.GenerateReport_Button.AutoSize = True
		Me.GenerateReport_Button.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
		Me.GenerateReport_Button.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!)
		Me.GenerateReport_Button.Location = New System.Drawing.Point(858, 34)
		Me.GenerateReport_Button.Name = "GenerateReport_Button"
		Me.GenerateReport_Button.Size = New System.Drawing.Size(140, 30)
		Me.GenerateReport_Button.TabIndex = 44
		Me.GenerateReport_Button.Text = "Generate Report"
		Me.GenerateReport_Button.UseVisualStyleBackColor = True
		'
		'Label1
		'
		Me.Label1.AutoSize = True
		Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label1.Location = New System.Drawing.Point(288, 39)
		Me.Label1.Name = "Label1"
		Me.Label1.Size = New System.Drawing.Size(65, 20)
		Me.Label1.TabIndex = 40
		Me.Label1.Text = "Board 1"
		'
		'CB_Board1
		'
		Me.CB_Board1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
		Me.CB_Board1.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.CB_Board1.FormattingEnabled = True
		Me.CB_Board1.Location = New System.Drawing.Point(359, 35)
		Me.CB_Board1.Name = "CB_Board1"
		Me.CB_Board1.Size = New System.Drawing.Size(208, 28)
		Me.CB_Board1.TabIndex = 41
		'
		'Close_Button
		'
		Me.Close_Button.AutoSize = True
		Me.Close_Button.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
		Me.Close_Button.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Close_Button.Location = New System.Drawing.Point(1119, 34)
		Me.Close_Button.Name = "Close_Button"
		Me.Close_Button.Size = New System.Drawing.Size(59, 30)
		Me.Close_Button.TabIndex = 46
		Me.Close_Button.Text = "Close"
		Me.Close_Button.UseVisualStyleBackColor = True
		'
		'CompareRevisions
		'
		Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.ClientSize = New System.Drawing.Size(1187, 581)
		Me.Controls.Add(Me.Label3)
		Me.Controls.Add(Me.DGV_Board2_Quantity)
		Me.Controls.Add(Me.DGV_Board1_Quantity)
		Me.Controls.Add(Me.L_Board1Quantities)
		Me.Controls.Add(Me.L_Board2Quantities)
		Me.Controls.Add(Me.Label8)
		Me.Controls.Add(Me.Label4)
		Me.Controls.Add(Me.Label9)
		Me.Controls.Add(Me.CB_Board2)
		Me.Controls.Add(Me.L_Title)
		Me.Controls.Add(Me.Label2)
		Me.Controls.Add(Me.CB_Source)
		Me.Controls.Add(Me.Excel_Button)
		Me.Controls.Add(Me.GenerateReport_Button)
		Me.Controls.Add(Me.Label1)
		Me.Controls.Add(Me.CB_Board1)
		Me.Controls.Add(Me.Close_Button)
		Me.Name = "CompareRevisions"
		Me.Text = "Compare Revisions"
		CType(Me.DGV_Board2_Quantity, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.DGV_Board1_Quantity, System.ComponentModel.ISupportInitialize).EndInit()
		Me.ResumeLayout(False)
		Me.PerformLayout()

	End Sub

	Friend WithEvents Label3 As Label
	Friend WithEvents DGV_Board2_Quantity As DataGridView
	Friend WithEvents DGV_Board1_Quantity As DataGridView
	Friend WithEvents L_Board1Quantities As Label
	Friend WithEvents L_Board2Quantities As Label
	Friend WithEvents Label8 As Label
	Friend WithEvents Label4 As Label
	Friend WithEvents Label9 As Label
	Friend WithEvents CB_Board2 As ComboBox
	Friend WithEvents L_Title As Label
	Friend WithEvents Label2 As Label
	Friend WithEvents CB_Source As ComboBox
	Friend WithEvents Excel_Button As Button
	Friend WithEvents GenerateReport_Button As Button
	Friend WithEvents Label1 As Label
	Friend WithEvents CB_Board1 As ComboBox
	Friend WithEvents Close_Button As Button
End Class
