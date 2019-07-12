'-----------------------------------------------------------------------------------------------------------------------------------------
' Module: PartUsage.vb
'
' Description: Creates a report that shows the use of parts over all of the boards that the user selects. 
'
' Buttons:
'	Refresh Mags = Looks in the directory that ALPHA exports to for a .mag file and compares it to the file in our release directory. If
'		the file is newer, it is then imported into the database for the user to use for the comparisions. This does not refresh the
'		report.
'
' Checkboxes:
'	All Vendors = Include the 2nd and 3rd Vendors and their MPN in the report.
'	Cost = Include the cost of a single part of the item.
'	Mag Info = Include any Magazine information if available.
'
' Special Keys:
'	delete = If we have an item highlighted in our chosenBoards listbox, then we remove it from the listbox.
'-----------------------------------------------------------------------------------------------------------------------------------------
Imports System.Data.SqlClient

Public Class PartUsage

	Dim DataTable_usage As DataTable

	Dim QBItemlist As DataTable
	Dim PCADBOMlist As DataTable
	Dim ALPHAMaglist As DataTable

	Private Sub PartUsage_Load() Handles MyBase.Load
		GetBoards()

		KeyPreview = True
	End Sub

	Private Sub GetBoards()
		Dim ProductNames As New DataTable()

		Dim myCmd As New SqlCommand("SELECT Distinct([" & DB_HEADER_BOARD_NAME & "]) FROM " & TABLE_PCADBOM & " ORDER BY [" & DB_HEADER_BOARD_NAME & "]", myConn)
		ProductNames.Load(myCmd.ExecuteReader)

		For Each dr As DataRow In ProductNames.Rows
			Boards_ListBox.Items.Add(dr(DB_HEADER_BOARD_NAME))
		Next
	End Sub

	Private Sub B_GenerateReport_Click() Handles B_GenerateReport.Click
		Cursor = Cursors.WaitCursor
		If LOGDATA = True Then
			Try
				If ChangeCheck(True) = True Then
					GenerateReport()
				End If
			Catch ex As Exception
				UnhandledExceptionMessage(ex)
			End Try
		Else
			If ChangeCheck(True) = True Then
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

		setupTable()

		QBItemlist = New DataTable
		Dim myCmd As New SqlCommand("SELECT * FROM " & TABLE_QB_ITEMS, myConn)
		QBItemlist.Load(myCmd.ExecuteReader())

		PCADBOMlist = New DataTable
		myCmd.CommandText = "SELECT * FROM " & TABLE_PCADBOM
		PCADBOMlist.Load(myCmd.ExecuteReader())

		ALPHAMaglist = New DataTable
		myCmd.CommandText = "SELECT * FROM " & TABLE_MAGAZINE_DATA
		ALPHAMaglist.Load(myCmd.ExecuteReader())

		Dim rownumber As Integer = 0

		Dim listOfBoards As New List(Of String)

		If ChosenBoards_ListBox.Items.Count <> 0 Then
			For Each item In ChosenBoards_ListBox.Items
				listOfBoards.Add(item)
			Next
		Else
			For Each item In Boards_ListBox.Items
				listOfBoards.Add(item)
			Next
		End If

		For Each item In listOfBoards
			'Grab the full BOM of the board that we have selected.
			Dim PCADdrs() As DataRow = PCADBOMlist.Select("[" & DB_HEADER_BOARD_NAME & "] = '" & item & "'")
			Dim ItemList As New List(Of String)

			For Each row As DataRow In PCADdrs
				'Check to see if we are dealing with a PCB or an item that we already added during this run through.
				If row(DB_HEADER_ITEM_PREFIX).ToString.Contains(PREFIX_PCB) = True Or ItemList.Contains(row(DB_HEADER_ITEM_NUMBER)) = True Then
					Continue For
				End If
				ItemList.Add(row(DB_HEADER_ITEM_NUMBER))

				'Check to see if we have added this item to the dataTable yet.
				Dim searchDR() As DataRow = DataTable_usage.Select("[" & DB_HEADER_ITEM_NUMBER & "] = '" & row(DB_HEADER_ITEM_PREFIX) & ":" & row(DB_HEADER_ITEM_NUMBER) & "'")

				If searchDR.Length = 0 Then
					'This is a new item and we need to add it to the table.
					DataTable_usage.Rows.Add()

					'Grab all of the information for this item.
					Dim QBdrs() As DataRow = QBItemlist.Select("[" & DB_HEADER_ITEM_NUMBER & "] = '" & row(DB_HEADER_ITEM_NUMBER) & "'")
					Dim magDR() As DataRow = ALPHAMaglist.Select("[" & DB_HEADER_ITEM_NUMBER & "] ='" & row(DB_HEADER_ITEM_NUMBER) & "'")

					If QBdrs.Length <> 0 Then
						'The item has been found in the database.
						DataTable_usage.Rows(rownumber)(DB_HEADER_PROCESS) = row(DB_HEADER_PROCESS)
						DataTable_usage.Rows(rownumber)(DB_HEADER_ITEM_NUMBER) = row(DB_HEADER_ITEM_PREFIX) & ":" & row(DB_HEADER_ITEM_NUMBER)
						DataTable_usage.Rows(rownumber)(DB_HEADER_DESCRIPTION) = row(DB_HEADER_DESCRIPTION)
						DataTable_usage.Rows(rownumber)(DB_HEADER_VENDOR) = QBdrs(0)(DB_HEADER_VENDOR)
						DataTable_usage.Rows(rownumber)(DB_HEADER_MPN) = QBdrs(0)(DB_HEADER_MPN)
						DataTable_usage.Rows(rownumber)(DB_HEADER_VENDOR2) = QBdrs(0)(DB_HEADER_VENDOR2)
						DataTable_usage.Rows(rownumber)(DB_HEADER_MPN2) = QBdrs(0)(DB_HEADER_MPN2)
						DataTable_usage.Rows(rownumber)(DB_HEADER_VENDOR3) = QBdrs(0)(DB_HEADER_VENDOR3)
						DataTable_usage.Rows(rownumber)(DB_HEADER_MPN3) = QBdrs(0)(DB_HEADER_MPN3)
						DataTable_usage.Rows(rownumber)(item) = PCADBOMlist.Select("[" & DB_HEADER_BOARD_NAME & "] = '" & item & "' AND [" & DB_HEADER_ITEM_NUMBER & "] = '" & row(DB_HEADER_ITEM_NUMBER) & "'").Length
						DataTable_usage.Rows(rownumber)(HEADER_TOTAL) = DataTable_usage.Rows(rownumber)(item)
						DataTable_usage.Rows(rownumber)(DB_HEADER_COST) = QBdrs(0)(DB_HEADER_COST)

						'If we have mag data then we need to add it to the table as well.
						If magDR.Length <> 0 Then
							DataTable_usage.Rows(rownumber)(DB_HEADER_NAME) = magDR(0)(DB_HEADER_NAME)
							DataTable_usage.Rows(rownumber)(DB_HEADER_SLOT_NUMBER) = magDR(0)(DB_HEADER_SLOT_NUMBER)
							DataTable_usage.Rows(rownumber)(DB_HEADER_FEEDER_NUMBER) = magDR(0)(DB_HEADER_FEEDER_NUMBER)
						End If

						rownumber += 1
					Else
						'This item is not in the database.
						DataTable_usage.Rows.Add()
						DataTable_usage.Rows(rownumber)(DB_HEADER_PROCESS) = row(DB_HEADER_PROCESS)
						DataTable_usage.Rows(rownumber)(DB_HEADER_ITEM_NUMBER) = row(DB_HEADER_ITEM_PREFIX) & ":" & row(DB_HEADER_ITEM_NUMBER)
						DataTable_usage.Rows(rownumber)(DB_HEADER_DESCRIPTION) = NOT_IN_DATABASE
						DataTable_usage.Rows(rownumber)(item) = PCADBOMlist.Select("[" & DB_HEADER_BOARD_NAME & "] = '" & item & "' AND [" & DB_HEADER_ITEM_NUMBER & "] = '" & row(DB_HEADER_ITEM_NUMBER) & "'").Length
						DataTable_usage.Rows(rownumber)(HEADER_TOTAL) = DataTable_usage.Rows(rownumber)(item)
						rownumber += 1
					End If
				Else
					'This item is already in the table and just need to add the count for the board.
					'If searchDR(0)(DB_HEADER_DESCRIPTION).ToString.Contains(NOT_IN_DATABASE) = False Then
					searchDR(0)(item) = PCADBOMlist.Select("[" & DB_HEADER_BOARD_NAME & "] = '" & item & "' AND [" & DB_HEADER_ITEM_NUMBER & "] = '" & row(DB_HEADER_ITEM_NUMBER) & "'").Length
					searchDR(0)(HEADER_TOTAL) = CInt(searchDR(0)(HEADER_TOTAL).ToString) + CInt(searchDR(0)(item).ToString)
					'End If
				End If
			Next
		Next

		'Check to see if we want to have the extra columns in our report.
		If AllVendors_CheckBox.Checked = False Then
			DataTable_usage.Columns.Remove(DB_HEADER_MPN2)
			DataTable_usage.Columns.Remove(DB_HEADER_MPN3)
			DataTable_usage.Columns.Remove(DB_HEADER_VENDOR2)
			DataTable_usage.Columns.Remove(DB_HEADER_VENDOR3)
		End If

		If Cost_CheckBox.Checked = False Then
			DataTable_usage.Columns.Remove(DB_HEADER_COST)
		End If

		If MagInfo_CheckBox.Checked = False Then
			DataTable_usage.Columns.Remove(DB_HEADER_NAME)
			DataTable_usage.Columns.Remove(DB_HEADER_FEEDER_NUMBER)
			DataTable_usage.Columns.Remove(DB_HEADER_SLOT_NUMBER)
		End If

		'Sort our list to show SMT first, Hand Flow next, and then Post Assembly last
		Dim tempTable As DataTable = DataTable_usage.Copy

		DataTable_usage.Rows.Clear()

		Dim drs As DataRow() = tempTable.Select("[" & DB_HEADER_PROCESS & "] like '%" & PROCESS_SMT & "%'", DB_HEADER_PROCESS & " ASC, " & DB_HEADER_ITEM_NUMBER & " ASC")
		For Each dr In drs
			Dim row As DataRow
			row = DataTable_usage.NewRow
			row.ItemArray = dr.ItemArray
			DataTable_usage.Rows.Add(row)
		Next

		drs = tempTable.Select("[" & DB_HEADER_PROCESS & "] = '" & PROCESS_HANDFLOW & "'", DB_HEADER_ITEM_NUMBER)
		For Each dr In drs
			Dim row As DataRow
			row = DataTable_usage.NewRow
			row.ItemArray = dr.ItemArray
			DataTable_usage.Rows.Add(row)
		Next

		drs = tempTable.Select("[" & DB_HEADER_PROCESS & "] = '" & PROCESS_POSTASSEMBLY & "'", DB_HEADER_ITEM_NUMBER)
		For Each dr In drs
			Dim row As DataRow
			row = DataTable_usage.NewRow
			row.ItemArray = dr.ItemArray
			DataTable_usage.Rows.Add(row)
		Next

		DGV_PartUsage.DataSource = Nothing
		DataTable_usage = DataTable_usage.DefaultView.ToTable
		DGV_PartUsage.DataSource = DataTable_usage
		DGV_PartUsage.AutoResizeColumns(DataGridViewAutoSizeColumnMode.AllCells)
		DGV_PartUsage.Columns(1).Frozen = True
		Excel_Button.Enabled = True
	End Sub

	Private Sub setupTable()
		DataTable_usage = New DataTable
		DataTable_usage.Columns.Add(DB_HEADER_PROCESS)
		DataTable_usage.Columns.Add(DB_HEADER_ITEM_NUMBER)
		DataTable_usage.Columns.Add(DB_HEADER_DESCRIPTION)
		DataTable_usage.Columns.Add(DB_HEADER_VENDOR)
		DataTable_usage.Columns.Add(DB_HEADER_MPN)

		'Optional Columns
		DataTable_usage.Columns.Add(DB_HEADER_VENDOR2)
		DataTable_usage.Columns.Add(DB_HEADER_MPN2)
		DataTable_usage.Columns.Add(DB_HEADER_VENDOR3)
		DataTable_usage.Columns.Add(DB_HEADER_MPN3)

		'For each board that we want to include, we need to add a column for it.
		If ChosenBoards_ListBox.Items.Count = 0 Then
			For Each board In Boards_ListBox.Items
				DataTable_usage.Columns.Add(board)
			Next
		Else
			For Each board In ChosenBoards_ListBox.Items
				DataTable_usage.Columns.Add(board)
			Next
		End If

		DataTable_usage.Columns.Add(HEADER_TOTAL)

		'Optional Columns
		DataTable_usage.Columns.Add(DB_HEADER_COST)

		DataTable_usage.Columns.Add(DB_HEADER_NAME)
		DataTable_usage.Columns.Add(DB_HEADER_SLOT_NUMBER)
		DataTable_usage.Columns.Add(DB_HEADER_FEEDER_NUMBER)
	End Sub

	Private Sub Excel_Button_Click() Handles Excel_Button.Click
		Dim report As New GenerateReport
		report.GeneratePartUsageReport(DataTable_usage)
	End Sub

	Private Sub B_Close_Click() Handles B_Close.Click
		Close()
	End Sub

	Private Sub MyBase_KeyDown(ByVal sender As Object, ByVal e As KeyEventArgs) Handles MyBase.KeyDown
		'If we press the Delete key, call our remove function.
		If e.KeyCode.Equals(Keys.Delete) Then
			Call Remove_Button_Click()
		End If
	End Sub

#Region "Dual Listbox Methods"
	Private Sub Products_ListBox_DoubleClick() Handles Boards_ListBox.DoubleClick
		'Treat the double click as if we are adding the item that we are double clicking on.
		Call Add_Button_Click()
	End Sub

	Private Sub Add_Button_Click() Handles Add_Button.Click
		'Add to our Build List using whatever board is selected in our product list only if we have not added it already.
		If ChosenBoards_ListBox.Items.Contains(Boards_ListBox.SelectedItem) = False Then
			ChosenBoards_ListBox.Items.Add(Boards_ListBox.SelectedItem)
		End If

		'We have changed the situation so disable the excel button so the user does not create under misinformation.
		Excel_Button.Enabled = False
	End Sub

	Private Sub Clear_Button_Click() Handles Clear_Button.Click
		'Clears the whole list of boards that we want to see the options for.
		ChosenBoards_ListBox.Items.Clear()

		'We have changed the situation so disable the excel button so the user does not create under misinformation.
		Excel_Button.Enabled = False
	End Sub

	Private Sub Remove_Button_Click() Handles Remove_Button.Click
		'Remove the selected board from our list of boards that we want to see the options for.
		If ChosenBoards_ListBox.SelectedItems.Count <> 0 Then
			ChosenBoards_ListBox.Items.Remove(ChosenBoards_ListBox.SelectedItem)

			'We have changed the situation so disable the excel button so the user does not create under misinformation.
			Excel_Button.Enabled = False
		End If
	End Sub
#End Region

	Private Sub DGV_QB_Items_RowPostPaint(ByVal sender As Object, ByVal e As DataGridViewRowPostPaintEventArgs) Handles DGV_PartUsage.RowPostPaint
		'Go through each row of the DGV and add the row number to the row header.
		Using b As SolidBrush = New SolidBrush(DGV_PartUsage.RowHeadersDefaultCellStyle.ForeColor)
			e.Graphics.DrawString(e.RowIndex + 1, DGV_PartUsage.DefaultCellStyle.Font, b, e.RowBounds.Location.X + 10, e.RowBounds.Location.Y + 4)
		End Using
	End Sub

	Private Sub RefreshMags_Button_Click() Handles RefreshMags_Button.Click
		Cursor = Cursors.WaitCursor
		If LOGDATA = True Then
			Try
				RefreshMags()
			Catch ex As Exception
				UnhandledExceptionMessage(ex)
			End Try
		Else
			RefreshMags()
		End If
		Cursor = Cursors.Default

		'Turn the dirty bit off once we are done with our import.
		sqlapi.SetDirtyBit(0)
	End Sub

End Class