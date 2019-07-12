'-----------------------------------------------------------------------------------------------------------------------------------------
' Module: FindStockNumber.vb
'
' Description: Search through the PCAD BOMs for a specific string. Options to search through are all handled with the combo box.
'-----------------------------------------------------------------------------------------------------------------------------------------
Imports System.Data.SqlClient

Public Class FindItemNumber

	Dim DataTable_Results As DataTable
	Dim DataSet_Resutls As DataSet

	Dim List_Boards As List(Of String)

	Private Sub FindStockNumber_Load() Handles MyBase.Load
		SetupTable()

		'Get drop down items.
		GetDropDownItems(CB_Search)
		CB_Search.DropDownHeight = 200

		GetDropDownItems(CB_Search2)
		CB_Search2.DropDownHeight = 200
	End Sub

	Private Sub SetupTable()
		DataTable_Results = New DataTable
		DataTable_Results.Columns.Add(DB_HEADER_BOARD_NAME)
		DataTable_Results.Columns.Add(DB_HEADER_ITEM_PREFIX)
		DataTable_Results.Columns.Add(DB_HEADER_ITEM_NUMBER)
		DataTable_Results.Columns.Add(DB_HEADER_DESCRIPTION)
		DataTable_Results.Columns.Add(DB_HEADER_MPN)
		DataTable_Results.Columns.Add(DB_HEADER_VENDOR)
		DataTable_Results.Columns.Add(DB_HEADER_PROCESS)
	End Sub

	Private Sub GetDropDownItems(ByRef box As ComboBox)
		Dim index As Integer = 0

		For Each dc As DataColumn In DataTable_Results.Columns
			'Using the syntaxName list in relation to the columnNames
			box.Items.Add(dc.ColumnName)
		Next

		If box.Items.Count <> 0 Then
			box.SelectedIndex = 0
		End If

	End Sub

	Private Sub B_Search_Click() Handles B_Search.Click
		Cursor = Cursors.WaitCursor
		If LOGDATA = True Then
			Try
				GenerateReport()
			Catch ex As Exception
				UnhandledExceptionMessage(ex)
			End Try
		Else
			GenerateReport()
		End If
		Cursor = Cursors.Default
	End Sub

	Private Sub GenerateReport()
		Dim message As String = ""
		If sqlapi.CheckDirtyBit(message) = True Then
			MsgBox(message)
			Return
		End If

		SetupTable()

		'Create our query.
		Dim query As String = "SELECT * FROM " & TABLE_PCADBOM & " WHERE [" & CB_Search.Text & "] LIKE '%" & TB_Search.Text & "%'"

		If TB_Search2.Text.Length <> 0 Then
			query = query & " AND [" & CB_Search2.Text & "] LIKE '%" & TB_Search2.Text & "%'"
		End If

		query = query & " AND [" & DB_HEADER_REF_DES & "] != '" & REFERENCE_DESIGNATOR_OPTION & "' ORDER BY [" & DB_HEADER_BOARD_NAME & "]"

		SearchBoards(query)

		If DataTable_Results.Rows.Count = 0 Then
			DataTable_Results.Rows.Add("No Results were found.")
		End If

		DataSet_Resutls = New DataSet()
		DataSet_Resutls.Tables.Add(DataTable_Results)

		DGV_Results.DataSource = Nothing
		DGV_Results.DataSource = DataSet_Resutls.Tables(0)
		DGV_Results.AutoResizeColumns(DataGridViewAutoSizeColumnMode.AllCells)
	End Sub

	Private Sub SearchBoards(ByRef query As String)
		Dim myCmd As New SqlCommand(query, myConn)
		Dim PCADinfo As New DataTable
		PCADinfo.Load(myCmd.ExecuteReader)
		Dim numberOfBoards As Integer = PCADinfo.DefaultView.ToTable(True, DB_HEADER_BOARD_NAME).Rows.Count

		For Each dsRow As DataRow In PCADinfo.Rows
			'Check to see if we already added the Stock Number AND Board Name.
			If DataTable_Results.Select("[" & DB_HEADER_BOARD_NAME & "] = '" & dsRow(DB_HEADER_BOARD_NAME) & "' AND [" & DB_HEADER_ITEM_NUMBER & "] = '" & dsRow(DB_HEADER_ITEM_PREFIX) & ":" & dsRow(DB_HEADER_ITEM_NUMBER) & "'").Length = 0 Then

				'Check to see if we have a prefix or not to include in the Stock Number.
				If dsRow(DB_HEADER_ITEM_PREFIX).ToString.Length = 0 Then
					'No prefix
					DataTable_Results.Rows.Add(dsRow(DB_HEADER_BOARD_NAME), dsRow(DB_HEADER_ITEM_PREFIX), dsRow(DB_HEADER_ITEM_NUMBER), dsRow(DB_HEADER_DESCRIPTION), dsRow(DB_HEADER_MPN), dsRow(DB_HEADER_VENDOR), dsRow(DB_HEADER_PROCESS))
				Else
					'Prefix
					DataTable_Results.Rows.Add(dsRow(DB_HEADER_BOARD_NAME), dsRow(DB_HEADER_ITEM_PREFIX), dsRow(DB_HEADER_ITEM_PREFIX) & ":" & dsRow(DB_HEADER_ITEM_NUMBER), dsRow(DB_HEADER_DESCRIPTION), dsRow(DB_HEADER_MPN), dsRow(DB_HEADER_VENDOR), dsRow(DB_HEADER_PROCESS))
				End If
			End If
		Next

		'Display number of boards and how many were found to the user.
		L_Found.Text = "Found in " & numberOfBoards & " Board(s). Results: " & DataTable_Results.Rows.Count
	End Sub

	Private Sub B_Close_Click() Handles B_Close.Click
		Close()
	End Sub

End Class