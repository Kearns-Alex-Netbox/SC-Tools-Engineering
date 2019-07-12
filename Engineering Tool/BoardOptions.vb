'-----------------------------------------------------------------------------------------------------------------------------------------
' Module: BoardOptions.vb
'
' Description: Select boards with options to see the description of each option
'
' Special Keys:
'	delete = If we have an item highlighted in our chosenBoards listbox, then we remove it from the listbox.
'-----------------------------------------------------------------------------------------------------------------------------------------
Imports System.Data.SqlClient

Public Class BoardOptions

	Dim boardOptions_Datatable As DataTable
	Dim PCADBOM As DataTable

	Private Sub BoardOptions_Load(ByVal sender As Object, ByVal e As EventArgs) Handles MyBase.Load
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

	Private Sub GenerateReport_Button_Click(ByVal sender As Object, ByVal e As EventArgs) Handles GenerateReport_Button.Click
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
		'Check to see if we are doing an import at the moment.
		Dim message As String = ""
		If sqlapi.CheckDirtyBit(message) = True Then
			MsgBox(message)
			Return
		End If

		setupTable()

		PCADBOM = New DataTable
		Dim myCmd As New SqlCommand("SELECT * FROM " & TABLE_PCADBOM & " WHERE [" & DB_HEADER_REF_DES & "] LIKE '" & REFERENCE_DESIGNATOR_OPTION & "%'", myConn)
		PCADBOM.Load(myCmd.ExecuteReader())

		'If we did not choose any boards, then do the report on all of them.
		'Else, go through each of the boards that we added to our list.
		If ChosenBoards_ListBox.Items.Count = 0 Then
			For Each item In Boards_ListBox.Items
				AddBoardDescription(item)
			Next
		Else
			For Each item In ChosenBoards_ListBox.Items
				AddBoardDescription(item)
			Next
		End If

		'Populate our DGV by removing the old source, adding the new source, and autosize the cells.
		BoardOptions_DGV.DataSource = Nothing
		BoardOptions_DGV.DataSource = boardOptions_Datatable
		BoardOptions_DGV.AutoResizeColumns(DataGridViewAutoSizeColumnMode.AllCells)

		'Allow us to create the results to Excel if we want to.
		Excel_Button.Enabled = True
	End Sub

	Private Sub setupTable()
		boardOptions_Datatable = New DataTable
		boardOptions_Datatable.Columns.Add(DB_HEADER_BOARD_NAME, GetType(String))
		boardOptions_Datatable.Columns.Add(DB_HEADER_DESCRIPTION, GetType(String))
	End Sub

	Private Sub AddBoardDescription(ByRef boardName As String)
		'Using the passed in board name, select the 'ZD' Reference Designator.
		Dim drs() As DataRow = PCADBOM.Select("[" & DB_HEADER_BOARD_NAME & "] = '" & boardName & "'")

		If drs.Length <> 0 Then
			boardOptions_Datatable.Rows.Add(boardName, drs(0)(DB_HEADER_DESCRIPTION))
		Else
			boardOptions_Datatable.Rows.Add(boardName, "ZD Reference Designator for this board does not exist.")
		End If

		'Add a space between our results.
		boardOptions_Datatable.Rows.Add()
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

	Private Sub Excel_Button_Click() Handles Excel_Button.Click
		Dim report As New GenerateReport
		report.GenerateBoardOptionReport(boardOptions_Datatable)
	End Sub

	Private Sub MyBase_KeyDown(ByVal sender As Object, ByVal e As KeyEventArgs) Handles MyBase.KeyDown
		'If we press the Delete key, call our remove function.
		If e.KeyCode.Equals(Keys.Delete) Then
			Call Remove_Button_Click()
		End If
	End Sub

	Private Sub Close_Button_Click() Handles Close_Button.Click
		Close()
	End Sub

End Class