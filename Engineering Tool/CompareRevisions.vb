'-----------------------------------------------------------------------------------------------------------------------------------------
' Module: CompareRevisions.vb
'
' Description: Compares revisions of two different boards. Board information can be sourced from PCAD, QuickBooks or Alpha
'-----------------------------------------------------------------------------------------------------------------------------------------
Imports System.Data.SqlClient

Public Class CompareRevisions

	Dim board1Quantity_DataTable As New DataTable
	Dim board2Quantity_DataTable As New DataTable

	Dim formLoaded As Boolean = False

	Private Sub CompareRevisions_Load() Handles MyBase.Load
		'Populate drop-down.
		GetSourceDropDownItems(CB_Source)
		CB_Source.DropDownHeight = 200

		GetBoardDropDownItems(CB_Board1)
		CB_Board1.DropDownHeight = 200

		GetBoardDropDownItems(CB_Board2)
		CB_Board2.DropDownHeight = 200

		formLoaded = True
		Me_Resize()
	End Sub

	Private Sub GenerateReport_Button_Click() Handles GenerateReport_Button.Click
		Cursor = Cursors.WaitCursor
		If LOGDATA = True Then
			Try
				If CB_Source.Text = TABLE_QBBOM Then
					If ChangeCheck(True) = True Then
						GenerateReport()
					End If
				Else
					GenerateReport()
				End If

			Catch ex As Exception
				UnhandledExceptionMessage(ex)
			End Try
		Else
			If CB_Source.Text = TABLE_QBBOM Then
				If ChangeCheck(True) = True Then
					GenerateReport()
				End If
			Else
				GenerateReport()
			End If
		End If
		Cursor = Cursors.Default
	End Sub

	Private Sub GenerateReport()
		Dim message As String = ""
		If sqlapi.CheckDirtyBit(message) = True Then
			MsgBox(message)
			Return
		End If

		'Check to make sure we are comparing two different board revisions.
		If CB_Board1.Text = CB_Board2.Text Then
			MsgBox("Please select two different boards to compare.")
			Return
		End If

		SetupQuantityTalbes()
		L_Title.Text = CB_Board1.Text & "   <->   " & CB_Board2.Text
		L_Title.Left = ClientSize.Width \ 2 - L_Title.Width \ 2
		L_Board1Quantities.Text = CB_Board1.Text & " Quantity"
		L_Board2Quantities.Text = CB_Board2.Text & " Quantity"

		'Depending on what source we choose, call the correct function.
		Select Case CB_Source.Text
			Case TABLE_ALPHABOM
				CompareALPHA_PCAD(CB_Board1.Text, CB_Board2.Text, TABLE_ALPHABOM)
			Case TABLE_PCADBOM
				CompareALPHA_PCAD(CB_Board1.Text, CB_Board2.Text, TABLE_PCADBOM)
			Case TABLE_QBBOM
				CompareQB(CB_Board1.Text, CB_Board2.Text)
        End Select

		'Check to see if we have added anything to the dataTables. If not, then they are the same.
		If board1Quantity_DataTable.Rows.Count = 0 Then
			board1Quantity_DataTable.Rows.Add("There are no quantity disagreements")
		End If
		If board2Quantity_DataTable.Rows.Count = 0 Then
			board2Quantity_DataTable.Rows.Add("There are no quantity disagreements")
		End If

		FormatGrid()

		Excel_Button.Enabled = True
	End Sub

	Private Sub SetupQuantityTalbes()
		board1Quantity_DataTable = New DataTable
		board2Quantity_DataTable = New DataTable

		board1Quantity_DataTable.Columns.Add(DB_HEADER_ITEM_NUMBER, GetType(String))
		board1Quantity_DataTable.Columns.Add(CB_Board1.Text & " quantity", GetType(String))
		board1Quantity_DataTable.Columns.Add(CB_Board2.Text & " quantity", GetType(String))

		board2Quantity_DataTable.Columns.Add(DB_HEADER_ITEM_NUMBER, GetType(String))
		board2Quantity_DataTable.Columns.Add(CB_Board1.Text & " quantity", GetType(String))
		board2Quantity_DataTable.Columns.Add(CB_Board2.Text & " quantity", GetType(String))
	End Sub

	Private Sub CompareALPHA_PCAD(ByRef board1 As String, ByRef board2 As String, ByRef table As String)
		'Used to compare either ALPHA or PCAD revisions
		Try
			'Create and add our query into our data set for comparison.
			Dim da As SqlDataAdapter = New SqlDataAdapter("SELECT * FROM " & table & " WHERE [" & DB_HEADER_BOARD_NAME & "] = '" & board1 & "' ORDER BY [" & DB_HEADER_ITEM_NUMBER & "]", myConn)
			Dim ds As DataSet = New DataSet()
			da.Fill(ds, "Board 1 DATA")

			'Create and add our query into our data set for comparison.
			da = New SqlDataAdapter("SELECT * FROM " & table & " WHERE [" & DB_HEADER_BOARD_NAME & "] = '" & board2 & "' ORDER BY [" & DB_HEADER_ITEM_NUMBER & "]", myConn)
			'ds = New DataSet()
			da.Fill(ds, "Board 2 DATA")

			'---------------------------------'
			'Compare Board 2 against Board 1. '
			'---------------------------------'
			For Each dsrow As DataRow In ds.Tables("Board 1 DATA").Rows
				'Get our count for the first board.
				Dim board1Quantity As Integer = ds.Tables("Board 1 DATA").Select("[" & DB_HEADER_ITEM_NUMBER & "] = '" & dsrow(DB_HEADER_ITEM_NUMBER) & "'").Length

				'Get our count for the second board.
				Dim board2Quantity As Integer = ds.Tables("Board 2 DATA").Select("[" & DB_HEADER_ITEM_NUMBER & "] = '" & dsrow(DB_HEADER_ITEM_NUMBER) & "'").Length

				If board1Quantity <> board2Quantity Then
					If board1Quantity_DataTable.Select("[" & DB_HEADER_ITEM_NUMBER & "] = '" & dsrow(DB_HEADER_ITEM_NUMBER) & "'").Length = 0 Then
						board1Quantity_DataTable.Rows.Add(dsrow(DB_HEADER_ITEM_NUMBER), board1Quantity, board2Quantity)
					End If
				End If
			Next

			'---------------------------------'
			'Compare Board 1 against Board 2. '
			'---------------------------------'
			For Each dsRow As DataRow In ds.Tables("Board 2 DATA").Rows
				'Get our count for the first board.
				Dim board1Quantity As Integer = ds.Tables("Board 1 DATA").Select("[" & DB_HEADER_ITEM_NUMBER & "] = '" & dsRow(DB_HEADER_ITEM_NUMBER) & "'").Length

				'Get our count for the second board.
				Dim board2Quantity As Integer = ds.Tables("Board 2 DATA").Select("[" & DB_HEADER_ITEM_NUMBER & "] = '" & dsRow(DB_HEADER_ITEM_NUMBER) & "'").Length

				If board1Quantity <> board2Quantity Then
					If board2Quantity_DataTable.Select("[" & DB_HEADER_ITEM_NUMBER & "] = '" & dsRow(DB_HEADER_ITEM_NUMBER) & "'").Length = 0 Then
						board2Quantity_DataTable.Rows.Add(dsRow(DB_HEADER_ITEM_NUMBER), board1Quantity, board2Quantity)
					End If
				End If
			Next
		Catch ex As Exception
			MsgBox(ex.Message)
			Return
		End Try
	End Sub

	Private Sub CompareQB(ByRef board1 As String, ByRef board2 As String)
		'QB report between revisions.
		Try
			Dim da As SqlDataAdapter = New SqlDataAdapter("SELECT * FROM " & TABLE_QBBOM & " WHERE [" & DB_HEADER_NAME & "] = '" & board1 & "' ORDER BY [" & DB_HEADER_ITEM_NUMBER & "]", myConn)
			Dim ds As DataSet = New DataSet()
			da.Fill(ds, "Board 1 DATA")

			'Create and add our query into our data set for comparison.
			da = New SqlDataAdapter("SELECT * FROM " & TABLE_QBBOM & " WHERE [" & DB_HEADER_NAME & "] = '" & board2 & "' ORDER BY [" & DB_HEADER_ITEM_NUMBER & "]", myConn)
			da.Fill(ds, "Board 2 DATA")

			'---------------------------------'
			'Compare Board 2 against Board 1. '
			'---------------------------------'
			For Each dsrow As DataRow In ds.Tables("Board 1 DATA").Rows
				Dim drs1() As DataRow = ds.Tables("Board 1 DATA").Select("[" & DB_HEADER_ITEM_NUMBER & "] = '" & dsrow(DB_HEADER_ITEM_NUMBER) & "'")
				Dim drs2() As DataRow = ds.Tables("Board 2 DATA").Select("[" & DB_HEADER_ITEM_NUMBER & "] = '" & dsrow(DB_HEADER_ITEM_NUMBER) & "'")

				'Get our count for the first board.

				Dim board1Quantity As Integer = 0
				If drs1.Length <> 0 Then
					board1Quantity = drs1(0)(DB_HEADER_QUANTITY)
				End If

				'Get our count for the second board.
				Dim board2Quantity As Integer = 0
				If drs2.Length <> 0 Then
					board2Quantity = drs2(0)(DB_HEADER_QUANTITY)
				End If

				If board1Quantity <> board2Quantity Then
					If board1Quantity_DataTable.Select("[" & DB_HEADER_ITEM_NUMBER & "] = '" & dsrow(DB_HEADER_ITEM_NUMBER) & "'").Length = 0 Then
						board1Quantity_DataTable.Rows.Add(dsrow(DB_HEADER_ITEM_NUMBER), board1Quantity, board2Quantity)
					End If
				End If
			Next

			'---------------------------------'
			'Compare Board 1 against Board 2. '
			'---------------------------------'
			For Each dsRow As DataRow In ds.Tables("Board 2 DATA").Rows
				Dim drs1() As DataRow = ds.Tables("Board 1 DATA").Select("[" & DB_HEADER_ITEM_NUMBER & "] = '" & dsRow(DB_HEADER_ITEM_NUMBER) & "'")
				Dim drs2() As DataRow = ds.Tables("Board 2 DATA").Select("[" & DB_HEADER_ITEM_NUMBER & "] = '" & dsRow(DB_HEADER_ITEM_NUMBER) & "'")

				'Get our count for the first board.

				Dim board1Quantity As Integer = 0
				If drs1.Length <> 0 Then
					board1Quantity = drs1(0)(DB_HEADER_QUANTITY)
				End If

				'Get our count for the second board.
				Dim board2Quantity As Integer = 0
				If drs2.Length <> 0 Then
					board2Quantity = drs2(0)(DB_HEADER_QUANTITY)
				End If

				If board1Quantity <> board2Quantity Then
					If board2Quantity_DataTable.Select("[" & DB_HEADER_ITEM_NUMBER & "] = '" & dsRow(DB_HEADER_ITEM_NUMBER) & "'").Length = 0 Then
						board2Quantity_DataTable.Rows.Add(dsRow(DB_HEADER_ITEM_NUMBER), board1Quantity, board2Quantity)
					End If
				End If
			Next
		Catch ex As Exception
			MsgBox(ex.Message)
			Return
		End Try
	End Sub

	'Format our data table. The structure that we are going for should have the common parts together at the bottom
	'	and then removed/added parts at the end of the table.

	'Part#		old Board	new Board
	'Common1			5	4
	'Common2			3	8
	'Removed1			2	 
	'Removed2			6	 
	'Added1				 	2
	'Added2				 	6
	Private Sub FormatGrid()
		Dim temp_Datatable As New DataTable
		temp_Datatable = board1Quantity_DataTable.Clone

		'end1 variable is used to line up the second table with the end of the first table to make it easier to see
		'which parts were removed/added.
		Dim end1 As Integer = 0

		Try
			'First, look for the parts that still have numbers between both revisions.
			For Each dtrow In board1Quantity_DataTable.Rows
				If dtrow(2) <> 0 Then
					temp_Datatable.Rows.Add(dtrow(DB_HEADER_ITEM_NUMBER), dtrow(CB_Board1.Text & " quantity"), dtrow(CB_Board2.Text & " quantity"))
					end1 += 1
				End If
			Next

			'Second, look for the parts that switched to 0 between revisions.
			For Each dtrow In board1Quantity_DataTable.Rows
				If dtrow(2) = 0 Then
					temp_Datatable.Rows.Add(dtrow(DB_HEADER_ITEM_NUMBER), dtrow(CB_Board1.Text & " quantity"), dtrow(CB_Board2.Text & " quantity"))
					end1 += 1
				End If
			Next

			board1Quantity_DataTable = New DataTable
			board1Quantity_DataTable = temp_Datatable.Copy
		Catch ex As Exception

		End Try

		DGV_Board1_Quantity.DataSource = board1Quantity_DataTable

		temp_Datatable = New DataTable
		temp_Datatable = board2Quantity_DataTable.Clone
		Dim end2 As Integer = 0

		Try
			'First, look for the parts that still have numbers between both revisions.
			For Each dtrow In board2Quantity_DataTable.Rows
				If dtrow(1) <> 0 Then
					temp_Datatable.Rows.Add(dtrow(DB_HEADER_ITEM_NUMBER), dtrow(CB_Board1.Text & " quantity"), dtrow(CB_Board2.Text & " quantity"))
					end2 += 1
				End If
			Next

			'Add blank rows until we have accounted for all of the parts that were removed from the previous revision.
			'KAR: black or back?
			While end2 < end1
				temp_Datatable.Rows.Add()
				end2 += 1
			End While

			'Second, look for the prats that switched to 0 between revisions.
			For Each dtrow In board2Quantity_DataTable.Rows
				If dtrow(1) = 0 Then
					temp_Datatable.Rows.Add(dtrow(DB_HEADER_ITEM_NUMBER), dtrow(CB_Board1.Text & " quantity"), dtrow(CB_Board2.Text & " quantity"))
					end2 += 1
				End If
			Next

			board2Quantity_DataTable = New DataTable
			board2Quantity_DataTable = temp_Datatable.Copy
		Catch ex As Exception

		End Try

		DGV_Board2_Quantity.DataSource = board2Quantity_DataTable
	End Sub

	Private Sub Excel_Button_Click() Handles Excel_Button.Click
		Dim report As New GenerateReport()
		report.GenerateRevisionReport(CB_Source.Text, CB_Board1.Text, CB_Board2.Text, board1Quantity_DataTable, board2Quantity_DataTable)
	End Sub

	Private Sub Close_Button_Click() Handles Close_Button.Click
		Close()
	End Sub

	Private Sub CB_Source_SelectedValueChanged() Handles CB_Source.SelectedValueChanged
		If formLoaded = True Then
			Excel_Button.Enabled = False
			GetBoardDropDownItems(CB_Board1)
			GetBoardDropDownItems(CB_Board2)
		End If
	End Sub

	Private Sub CB_Board1_SelectedValueChanged() Handles CB_Board1.SelectedValueChanged, CB_Board2.SelectedValueChanged
		Excel_Button.Enabled = False
	End Sub

	Sub GetSourceDropDownItems(ByRef box As ComboBox)
		Dim BoardNames As New DataTable()

		Dim myCmd As New SqlCommand("SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE = 'Base Table' AND TABLE_SCHEMA = 'dbo' AND TABLE_NAME LIKE '%BOM'", myConn)

		BoardNames.Load(myCmd.ExecuteReader)

		For Each dr As DataRow In BoardNames.Rows
			Select Case dr("TABLE_NAME")
				Case TABLE_ALPHABOM
					box.Items.Add(TABLE_ALPHABOM)
				Case TABLE_PCADBOM
					box.Items.Add(TABLE_PCADBOM)
				Case TABLE_QBBOM
					box.Items.Add(TABLE_QBBOM)
			End Select
		Next

		If box.Items.Count <> 0 Then
			box.SelectedIndex = 0
		End If

	End Sub

	Sub GetBoardDropDownItems(ByRef box As ComboBox)
		box.Items.Clear()
		Dim BoardNames As New DataTable()
		Dim columnName As String = ""
		Dim extra As String = ""

		Select Case CB_Source.Text
			Case TABLE_ALPHABOM
				columnName = DB_HEADER_BOARD_NAME
			Case TABLE_PCADBOM
				columnName = DB_HEADER_BOARD_NAME
			Case TABLE_QBBOM
				columnName = DB_HEADER_NAME
				extra = " WHERE [" & DB_HEADER_NAME_PREFIX & "] != '" & PREFIX_FGS & "'"
		End Select

		Dim myCmd As New SqlCommand("SELECT Distinct([" & columnName & "]) FROM " & CB_Source.Text & extra & " ORDER BY [" & columnName & "]", myConn)

		BoardNames.Load(myCmd.ExecuteReader)

		For Each dr As DataRow In BoardNames.Rows
			box.Items.Add(dr(columnName))
		Next

		If box.Items.Count <> 0 Then
			box.SelectedIndex = 0
		End If

	End Sub

	Private Sub Me_Resize() Handles Me.Resize
		'Recalculate new column widths based on new size of window
		Dim newWidth As Integer = ClientSize.Width / 2
		Dim leftAndRightPadding As Integer = 16
		Dim topAndBottomPadding As Integer = 126

		L_Board2Quantities.Location = New Point(newWidth + 3, L_Board1Quantities.Location.Y)
		DGV_Board2_Quantity.Location = New Point((L_Board2Quantities.Location.X + 3), DGV_Board1_Quantity.Location.Y)
		DGV_Board2_Quantity.Width = newWidth - leftAndRightPadding
		DGV_Board2_Quantity.Height = ClientSize.Height - topAndBottomPadding

		DGV_Board1_Quantity.Width = newWidth - leftAndRightPadding
		DGV_Board1_Quantity.Height = ClientSize.Height - topAndBottomPadding
	End Sub

End Class