'-----------------------------------------------------------------------------------------------------------------------------------------
' Module: MasterControl.vb
'
' Description: Lou's special Process Release utilities
'   Tab 1 QuickBooks: Items: List of QB items with search ability
'	Tab 2 BOM Compare: Do BOM [from database Or from file] syntax check along with items And QB BOM checks before releasing
'   Tab 3 Release: Do directory comparisons to ensure all components exist before copying released files into release folder
'	Tab 4 PCAD Build: Takes the selected RELEASED BOM from PCAD and builds attempts to build the specified quantity.
'
' Tab 5 Buttons:
'	Compare QB Items = Compares the PCAD BOM that was selected with all of the items inside of QB. If an item is not found, then it will be
'		listed to the user.
'	Compare QB BOM = Compares the PCAD BOM that was selected with the QB BOM [if it is inside the database]. Differences will be displayed
'		to the user.
'	Processes = Opens the process window.
'	Reload Search = reloads the current file search. This allows the user to change the file and easily redo the search and compare.
'
' Tab 5 Checkboxes:
'	Show Only Differences = Do not show all of the items, only the ones that are different. Limited to 
'		Item Number and Item Prefix.
'	Include Manufacture and Part = ONLY APPLIES to [Compare QB Items] If [Show Only Differences is checked] is checked, then expand the 
'		compare to include vendor And MPN.
'
' Tab 5 Indicator Light: Shows what the results of the compare is.
'	Black = Not Applicable.
'	Green = There are no errors with the compare.
'	Red = There are errors that need to be fixed.
'
' Tab 5 Double-click Row Header: ONLY applies to a compare between PCAD/BOM file and QB BOM. Looks through the list of the opposite DGV to 
'		find the item. If it is found, then it will display to the user what the differences are if any.
'
' Tab 6 Buttons:
'	ALPHA [Source] = Makes the source location set to [\\Server1\Shares\Production\AlphaBackup] hard coded right now.
'	Check = Runs the check for the current source location. Allows the user to change a file(s) in the location and then re-run the check.
'	Back = Navigates the source folder location up one folder level.
'	Copy Release Files = Takes the files that are required for release and copies them over to the destination folder location. These
'		files include: all .BOM.CSV, .PCB, .PNP.CSV, .SCH, .SCH.PDF, and the Release Folder. EXCEPTION Only will be enabled if the source
'		location has passed the release check. Destination files/folders cannot be Read Only.
'	Build Options = Reads through the BOM file and gets all of the valid options and creates seperate files for each option with the 
'		correct name.
'	Delete [Source] = ONLY APPLIES to the ALPHA location. Deletes the selected file(s)/folder(s) that are selected.
'	ALPHA [Destination] = explorer window opens at the root of the released directory to allow the user to chose what ALPHA destination
'		they would like to copy files to.
'	Add = Takes the selected files from the source side and copies them over to the destination side.
'	Delete [Destination] = Deletes the selected file(s)/folder(s) that are selected. EXCEPTION Destination files/folders cannot be Read
'		Only.
'
' Tab 6 Checkboxes:
'	Read Only = Makes the Destination location files/folders Read Only.
'
' Tab 6 Indicator Lights: Shows what the results are from checking the destination to see if they are ready for release.
'	Black = Not Applicable.
'	Green = There are no errors with the check.
'	Red = There are errors that need to be fixed.
'-----------------------------------------------------------------------------------------------------------------------------------------
Imports System.Data.SqlClient
Imports System.IO
Imports Microsoft.VisualBasic.FileIO

Public Class MasterControl

#Region "Tab 1: QB Items Variables"
	Dim QBitems_da As New SqlDataAdapter
	Dim QBitems_ds As New DataSet
	Dim QBitems_myCmd = New SqlCommand("", myConn)

	Dim QBitems_scrollValue As Integer
	Dim QBitems_Command As String = "SELECT [" & DB_HEADER_ITEM_PREFIX & "]" &
										", [" & DB_HEADER_ITEM_PREFIX & "] + ':' + [" & DB_HEADER_ITEM_NUMBER & "] AS '" & DB_HEADER_ITEM_NUMBER & "'" &
										", [" & DB_HEADER_TYPE & "]" &
										", [" & DB_HEADER_DESCRIPTION & "]" &
										", [" & DB_HEADER_VENDOR & "]" &
										", [" & DB_HEADER_MPN & "]" &
										", [" & DB_HEADER_QUANTITY & "]" &
										", [" & DB_HEADER_COST & "]" &
										", [" & DB_HEADER_LEAD_TIME & "]" &
										", [" & DB_HEADER_MIN_ORDER_QTY & "]" &
										", [" & DB_HEADER_REORDER_QTY & "]" &
										", [" & DB_HEADER_VENDOR2 & "]" &
										", [" & DB_HEADER_MPN2 & "]" &
										", [" & DB_HEADER_VENDOR3 & "]" &
										", [" & DB_HEADER_MPN3 & "]" &
										" FROM " & TABLE_QB_ITEMS
	Dim QBitems_countCommand As String = "SELECT COUNT(*) FROM " & TABLE_QB_ITEMS
	Dim QBitems_entriesToShow As Integer = 250
	Dim QBitems_numberOfRecords As Integer
	Dim QBitems_sort As String = ""
	Dim QBitems_searchCommand As String = ""
	Dim QBitems_searchCountCommand As String = ""
	Dim QBitems_Freeze As Integer = 1
#End Region

#Region "Tab 2: BOM Compare Variables"
	Dim PCAD_BOM_da As New SqlDataAdapter
	Dim PCAD_BOM_ds As New DataSet
	Dim PCAD_BOM_myCmd = New SqlCommand("", myConn)

	Dim QB_BOM_da As New SqlDataAdapter
	Dim QB_BOM_ds As New DataSet
	Dim QB_BOM_myCmd = New SqlCommand("", myConn)

	Dim ALPHA_BOM_da As New SqlDataAdapter
	Dim ALPHA_BOM_ds As New DataSet
	Dim ALPHA_BOM_myCmd = New SqlCommand("", myConn)

	Dim fromPCADdatabase As Boolean = Nothing
	Dim fromSearch As Boolean = Nothing
	Dim fromCompareItems As Boolean = Nothing
	Dim fromCompareBOM As Boolean = Nothing
	Dim fromCompareALPHA As Boolean = Nothing
#End Region

#Region "Tab 3: Release Variables"
	'File Locations
	Dim ALPHA_BACKUP As String = "\\Server1\Shares\Production\AlphaBackup"
	Dim RELEASE_PCAD As String = My.Settings.ReleaseLocation & "\PCAD"
	Dim RELEASE As String = "\\Server1\EngineeringReleased\Boards"

	Dim DataTable_source As DataTable
	Dim DataTable_destination As DataTable

	Dim optionListSource As List(Of String)
	Dim optionListDestination As List(Of String)

	Dim fromTP_Compare As Boolean = False
#End Region

#Region "Tab 4: PCAD Build Variables"
	Dim numberToBuild As Integer = 0

	Dim build_ds As New DataSet

	'Styles used to help see results on the datagrid.
	Dim OUT_COLOR As Color = Color.FromArgb(255, 151, 163)
	Dim DATABASE_COLOR As Color = Color.Orange

	Const MICROMETTER_CONVERTER As Integer = 25.4

	Dim fu1X As Integer = 0
	Dim fu1Y As Integer = 0
	Dim fu2X As Integer = 0
	Dim fu2Y As Integer = 0
	Dim fu3X As Integer = 0
	Dim fu3Y As Integer = 0
	Dim boardRotation As Integer = 0
	Dim boardMirror As Integer = 0
	Dim fiducialName As String = "Netbox.10"

#End Region

	Dim myCmd As New SqlCommand("", myConn)

	Private Sub MasterControl_Load() Handles MyBase.Load
		Dim message As String = ""
		If sqlapi.CheckDirtyBit(message) = True Then
			MsgBox(message)
			Return
		End If

		'Inherited Sub that will center the opening form to screen.
		CenterToParent()

		Dim result As String = ""

		Try
			'Tab Page 1 - QuickBooks
			sqlapi.GetNumberOfRecords(QBitems_myCmd, QBitems_countCommand, QBitems_numberOfRecords, result)
			L_QB_Results.Text = "Number of results: " & QBitems_numberOfRecords

			QBitems_myCmd.CommandText = QBitems_Command & " ORDER BY [" & DB_HEADER_ITEM_PREFIX & "] + ':' + [" & DB_HEADER_ITEM_NUMBER & "]"
			QBitems_da = New SqlDataAdapter(QBitems_myCmd)
			QBitems_ds = New DataSet()

			sqlapi.RetriveData(QBitems_Freeze, QBitems_da, QBitems_ds, DGV_QB_Items, QBitems_scrollValue, QBitems_entriesToShow, QBitems_numberOfRecords,
							   B_QB_Next, B_QB_Last, B_QB_First, B_QB_Previous)

			'Get Drop Down Items.
			GetColumnDropDownItems(CB_QB_Sort, QBitems_ds)
			GetColumnDropDownItems(CB_QB_Search, QBitems_ds)
			GetColumnDropDownItems(CB_QB_Search2, QBitems_ds)

			CB_QB_Operand1.SelectedIndex = 0
			CB_QB_Operand2.SelectedIndex = 0

			'TAB Page 2
			'Get Drop Down Items.
			GetBoardDropDownItems(CB_Boards)

			'Tab Page 3
			TB_FilePath.Text = My.Settings.BOMFilePath

			'Tab Page 4
			'Get Drop Down Items.
			GetBoardDropDownItems(BoardBuild_ComboBox)
			BuildBoardSearch_TB.Text = My.Settings.BOMFilePath

			TB_ImportIndicator.BackColor = Color.LightGreen
			KeyPreview = True
		Catch ex As Exception
			MsgBox(ex.Message)
		End Try
	End Sub

	Private Sub QBItems_Resize() Handles Me.Resize
		ResizeTables(TabControl1.SelectedTab)
	End Sub

	Public Sub ResizeTables(ByRef Tab As TabPage)
		'Recalculate new column widths based on new size of window
		Dim newWidth As Integer = ClientSize.Width / 2
		Dim leftAndRightPadding As Integer = 40
		Dim topAndBottomPadding As Integer = 5

		'Tab Page 2 - BOM
		If fromCompareBOM = True Or fromCompareItems = True Or fromCompareALPHA = True Then
			DGV_QB_BOM.Location = New Point(newWidth, 74)
			DGV_QB_BOM.Width = newWidth - leftAndRightPadding

			DGV_PCAD_BOM.Width = newWidth - leftAndRightPadding
		Else
			DGV_PCAD_BOM.Width = Tab.Width - leftAndRightPadding
		End If

		'Tab Page 3 - Release
		DGV_Destination.Location = New Point(newWidth + 2, 158)
		DGV_Destination.Width = newWidth - leftAndRightPadding

		DGV_Source.Width = newWidth - leftAndRightPadding

		L_Destination.Location = New Point(newWidth + 2, 5)

		TB_DestinationFolderPath.Location = New Point(newWidth + 2, 28)

		B_SearchDestination.Location = New Point(newWidth + 2, 60)

		B_ALPHAdestination.Location = New Point(B_SearchDestination.Location.X + 76, 60)

		TB_DestinationIndicatorLight.Location = New Point(newWidth + 2, 96)

		L_Release.Location = New Point(TB_DestinationIndicatorLight.Location.X + 26, 96)

		B_Add.Location = New Point(newWidth + 2, 122)

		B_DeleteDestination.Location = New Point(B_Add.Location.X + 54, 122)

		CkB_ReadOnly.Location = New Point(B_Add.Location.X + 126, 126)

		TabControl1.Refresh()
	End Sub

	Private Sub B_CreateExcel_Click() Handles B_CreateExcel.Click
		Dim report As New GenerateReport()

		'Depending on which tab is open will determine which report to create.
		Select Case TabControl1.SelectedTab.Name
			Case "TP_QB_items"
				Dim Temp_ds As New DataSet
				QBitems_da.Fill(Temp_ds, "TABLE")
				report.GenerateQB_itemslistReport(Temp_ds)
			Case "TP_BOM_compare"
				'First check to see if we have made a report.
				If PCAD_BOM_ds Is Nothing Then
					MsgBox("Please make a table first before you create to Excel.")
				End If

				'Second check to see if our report is a single or double DGV
				If fromPCADdatabase Then
					If fromCompareBOM = True Or fromCompareItems = True Then
						report.GenerateBOMCompareReport(PCAD_BOM_ds, QB_BOM_ds, CB_Boards.Text)
					Else
						report.GenerateBOMCompareReport(PCAD_BOM_ds, Nothing, CB_Boards.Text)
					End If
				ElseIf fromSearch Then
					If fromCompareBOM = True Or fromCompareItems = True Then
						report.GenerateBOMCompareReport(PCAD_BOM_ds, QB_BOM_ds, Path.GetFileName(My.Settings.BOMFilePath))
					Else
						report.GenerateBOMCompareReport(PCAD_BOM_ds, Nothing, Path.GetFileName(My.Settings.BOMFilePath))
					End If
				End If
			Case "TP_PCAD_Build"
				report.GenerateQB_itemslistReport(build_ds)
			Case Else
				MsgBox("Generating a report on this tab has not been coded yet.")
		End Select
	End Sub

	Private Sub B_UpdateQB_Click() Handles B_UpdateQB.Click
		'This check allows the user to reset the flag that tells us if an import is going on. This way if the program crashes, we have a way to recover.
		Dim message As String = ""
		Dim answer As DialogResult
		If sqlapi.CheckDirtyBit(message) = True Then
			answer = MessageBox.Show(message & vbNewLine & "Would you like to reset the flag??", "Reset?", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)
			If answer = Windows.Forms.DialogResult.Yes Then
				sqlapi.SetDirtyBit(0)
			End If
			Return
		End If

		'Check to see if we can connect to the QB Database
		Try
			_cn.Open()
		Catch ex As Exception
			MsgBox("Could not connect to the QuickBooks database. Please make sure QuickBooks is open.")
			Return
		End Try

		'Change our indicator that an import is going on.
		TB_ImportIndicator.BackColor = Color.Gray
		TB_ImportIndicator.Refresh()

		Cursor = Cursors.WaitCursor

		Dim myCmd = New SqlCommand("", myConn)

		'Set our dirty bit to prevent others from runing reports while the import is going on.
		sqlapi.SetDirtyBit(1)

		'Delete all of our data in our QB BOM and Items tables.
		myCmd.CommandText = "DELETE FROM " & TABLE_QBBOM
		myCmd.ExecuteNonQuery()
		myCmd.CommandText = "DELETE FROM " & TABLE_QB_ITEMS
		myCmd.ExecuteNonQuery()

		'Run our import for the BOM and Items.
		Dim immports As New ImportData(Nothing)
		immports.ImportQBItems()
		immports.ImportQBBOMS()

		'Update our lastUpdate time in the database.
		myCmd = New SqlCommand("UPDATE " & TABLE_UTILITIES & " SET [" & DB_HEADER_VALUE & "] = GETDATE() WHERE [" & DB_HEADER_NAME & "] = 'LastUpdate'", myConn)
		myCmd.ExecuteNonQuery()

		'Reset the dirty bit, close the connection, and change our indicators.
		sqlapi.SetDirtyBit(0)
		_cn.Close()
		TB_ImportIndicator.BackColor = Color.LightGreen
		Cursor = Cursors.Default
		MsgBox("Import complete.")
	End Sub

	Private Sub B_Close_Click() Handles B_Close.Click
		Close()
	End Sub

	Private Sub GetColumnDropDownItems(ByRef cb As ComboBox, ByRef ds As DataSet)
		For Each dc As DataColumn In ds.Tables(0).Columns
			cb.Items.Add(dc.ColumnName)
		Next

		If cb.Items.Count <> 0 Then
			cb.SelectedIndex = 0
		End If

		cb.DropDownHeight = 200
	End Sub

#Region "Tab 1: QB Items"
	Private Sub B_QB_First_Click() Handles B_QB_First.Click
		sqlapi.FirstPage(QBitems_scrollValue, QBitems_ds, QBitems_da, QBitems_entriesToShow)
		B_QB_First.Enabled = False
		B_QB_Previous.Enabled = False
		B_QB_Next.Enabled = True
		B_QB_Last.Enabled = True
	End Sub

	Private Sub B_QB_Previous_Click() Handles B_QB_Previous.Click
		sqlapi.PreviousPage(QBitems_scrollValue, QBitems_entriesToShow, QBitems_ds, QBitems_da, B_QB_First, B_QB_Previous)
		B_QB_Next.Enabled = True
		B_QB_Last.Enabled = True
	End Sub

	Private Sub B_QB_Next_Click() Handles B_QB_Next.Click
		sqlapi.NextPage(QBitems_scrollValue, QBitems_entriesToShow, QBitems_numberOfRecords, QBitems_ds, QBitems_da, B_QB_Next, B_QB_Last)
		B_QB_First.Enabled = True
		B_QB_Previous.Enabled = True
	End Sub

	Private Sub B_QB_Last_Click() Handles B_QB_Last.Click
		sqlapi.LastPage(QBitems_scrollValue, QBitems_entriesToShow, QBitems_numberOfRecords, QBitems_ds, QBitems_da)
		B_QB_First.Enabled = True
		B_QB_Previous.Enabled = True
		B_QB_Next.Enabled = False
		B_QB_Last.Enabled = False
	End Sub

	Private Sub B_QB_ListAll_Click() Handles B_QB_ListAll.Click
		Cursor = Cursors.WaitCursor
		If LOGDATA = True Then
			Try
				sqlapi.ListAll(1, QBitems_searchCommand, QBitems_searchCountCommand, QBitems_myCmd, QBitems_countCommand, QBitems_numberOfRecords, L_QB_Results,
							   QBitems_Command & " ORDER BY [" & DB_HEADER_ITEM_PREFIX & "] + ':' + [" & DB_HEADER_ITEM_NUMBER & "]", QBitems_ds, QBitems_da, DGV_QB_Items, QBitems_scrollValue, QBitems_entriesToShow, B_QB_Next,
							   B_QB_Last, B_QB_First, B_QB_Previous)
			Catch ex As Exception
				UnhandledExceptionMessage(ex)
			End Try
		Else
			sqlapi.ListAll(1, QBitems_searchCommand, QBitems_searchCountCommand, QBitems_myCmd, QBitems_countCommand, QBitems_numberOfRecords, L_QB_Results,
						   QBitems_Command & " ORDER BY [" & DB_HEADER_ITEM_PREFIX & "] + ':' + [" & DB_HEADER_ITEM_NUMBER & "]", QBitems_ds, QBitems_da, DGV_QB_Items, QBitems_scrollValue, QBitems_entriesToShow, B_QB_Next,
						   B_QB_Last, B_QB_First, B_QB_Previous)
		End If
		Cursor = Cursors.Default
	End Sub

	Private Sub B_QB_Search_Click() Handles B_QB_Search.Click
		Cursor = Cursors.WaitCursor
		If LOGDATA = True Then
			Try
				sqlapi.Search(QBitems_Freeze, TB_QB_Search, CB_QB_Operand1, QBitems_searchCommand, QBitems_Command, QBitems_searchCountCommand, QBitems_countCommand, TB_QB_Search2, CB_QB_Operand2,
							  CB_QB_Search, CB_QB_Search2, QBitems_myCmd, QBitems_ds, QBitems_da, DGV_QB_Items, QBitems_numberOfRecords, L_QB_Results, QBitems_scrollValue,
							  QBitems_entriesToShow, B_QB_Next, B_QB_Last, B_QB_First, B_QB_Previous, "[" & DB_HEADER_ITEM_PREFIX & "] + ':' + [" & DB_HEADER_ITEM_NUMBER & "]")
			Catch ex As Exception
				UnhandledExceptionMessage(ex)
			End Try
		Else
			sqlapi.Search(QBitems_Freeze, TB_QB_Search, CB_QB_Operand1, QBitems_searchCommand, QBitems_Command, QBitems_searchCountCommand, QBitems_countCommand, TB_QB_Search2, CB_QB_Operand2,
						  CB_QB_Search, CB_QB_Search2, QBitems_myCmd, QBitems_ds, QBitems_da, DGV_QB_Items, QBitems_numberOfRecords, L_QB_Results, QBitems_scrollValue,
						  QBitems_entriesToShow, B_QB_Next, B_QB_Last, B_QB_First, B_QB_Previous, "[" & DB_HEADER_ITEM_PREFIX & "] + ':' + [" & DB_HEADER_ITEM_NUMBER & "]")
		End If
		Cursor = Cursors.Default
	End Sub

	Private Sub B_QB_Sort_Click() Handles B_QB_Sort.Click
		Cursor = Cursors.WaitCursor
		If LOGDATA = True Then
			Try
				sqlapi.Sort(QBitems_Freeze, QBitems_searchCommand, QBitems_Command, CB_QB_Sort, RB_QB_AscendingSort, QBitems_myCmd, QBitems_ds, QBitems_da, DGV_QB_Items,
							QBitems_scrollValue, QBitems_entriesToShow, QBitems_numberOfRecords, B_QB_Next, B_QB_Last, B_QB_First, B_QB_Previous)
			Catch ex As Exception
				UnhandledExceptionMessage(ex)
			End Try
		Else
			sqlapi.Sort(QBitems_Freeze, QBitems_searchCommand, QBitems_Command, CB_QB_Sort, RB_QB_AscendingSort, QBitems_myCmd, QBitems_ds, QBitems_da, DGV_QB_Items,
						QBitems_scrollValue, QBitems_entriesToShow, QBitems_numberOfRecords, B_QB_Next, B_QB_Last, B_QB_First, B_QB_Previous)
		End If
		Cursor = Cursors.Default

	End Sub

	Private Sub TB_QB_Search_KeyDown(ByVal sender As Object, ByVal e As KeyEventArgs) Handles TB_QB_Search.KeyDown
		If e.KeyCode.Equals(Keys.Enter) Then
			Call B_QB_Search_Click()
		End If
	End Sub

	Private Sub TB_QB_Search2_KeyDown(ByVal sender As Object, ByVal e As KeyEventArgs) Handles TB_QB_Search2.KeyDown
		If e.KeyCode.Equals(Keys.Enter) Then
			Call B_QB_Search_Click()
		End If
	End Sub

	Private Sub CB_QB_Display_SelectedValueChanged() Handles CB_QB_Display.SelectedValueChanged
		QBitems_entriesToShow = CInt(CB_QB_Display.Text)
	End Sub

	Private Sub CB_QB_Search_Click() Handles CB_QB_Search.Click
		CB_QB_Search.SelectedIndex = 0
	End Sub

	Private Sub CB_QB_Search_DropDownClosed() Handles CB_QB_Search.DropDownClosed
		Dim newcmd As New SqlCommand("SELECT DATA_TYPE FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = '" & TABLE_QB_ITEMS & "' AND COLUMN_NAME = '" & CB_QB_Search.Text & "'", myConn)

		Dim type As String = newcmd.ExecuteScalar

		If type = "decimal" Or type = "int" Then
			CB_QB_Operand1.SelectedIndex = 3
		Else
			CB_QB_Operand1.SelectedIndex = 0
		End If
	End Sub

	Private Sub CB_QB_Search2_DropDownClosed() Handles CB_QB_Search2.DropDownClosed
		Dim newcmd As New SqlCommand("SELECT DATA_TYPE FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = '" & TABLE_QB_ITEMS & "' AND COLUMN_NAME = '" & CB_QB_Search2.Text & "'", myConn)

		Dim type As String = newcmd.ExecuteScalar

		If type = "decimal" Or type = "int" Then
			CB_QB_Operand2.SelectedIndex = 3
		Else
			CB_QB_Operand2.SelectedIndex = 0
		End If
	End Sub

	Private Sub CB_QB_Sort_Click() Handles CB_QB_Sort.Click
		CB_QB_Sort.SelectedIndex = 0
	End Sub

	Private Sub DGV_QB_Items_RowPostPaint(ByVal sender As Object, ByVal e As DataGridViewRowPostPaintEventArgs) Handles DGV_QB_Items.RowPostPaint
		'Go through each row of the DGV and add the row number to the row header.
		Using b As SolidBrush = New SolidBrush(DGV_QB_Items.RowHeadersDefaultCellStyle.ForeColor)
			e.Graphics.DrawString(e.RowIndex + 1 + QBitems_scrollValue, DGV_QB_Items.DefaultCellStyle.Font, b, e.RowBounds.Location.X + 10, e.RowBounds.Location.Y + 4)
		End Using
	End Sub
#End Region

#Region "Tab 2: BOM Compare"
	Private Sub B_CreateBOM_Click() Handles B_CreateBOM.Click
		Cursor = Cursors.WaitCursor
		If LOGDATA = True Then
			Try
				If ChangeCheck(True) = True Then
					CreateBOM()
				End If
			Catch ex As Exception
				UnhandledExceptionMessage(ex)
			End Try
		Else
			If ChangeCheck(True) = True Then
				CreateBOM()
			End If
		End If
		Cursor = Cursors.Default
	End Sub

	Private Sub CreateBOM()
		Dim message As String = ""
		If sqlapi.CheckDirtyBit(message) = True Then
			MsgBox(message)
			Return
		End If

		TB_SearchIndicator.BackColor = Color.Black
		PCAD_BOM_myCmd.CommandText = "SELECT [" & DB_HEADER_REF_DES & "], [" & DB_HEADER_DESCRIPTION & "], [" & DB_HEADER_VENDOR & "], [" & DB_HEADER_MPN & "], COALESCE(NULLIF([" & DB_HEADER_ITEM_PREFIX & "] + ':' + [" & DB_HEADER_ITEM_NUMBER & "],':'), [" & DB_HEADER_ITEM_NUMBER & "]) AS '" & DB_HEADER_ITEM_NUMBER & "', [" & DB_HEADER_PROCESS & "]  FROM " & TABLE_PCADBOM & " WHERE [" & DB_HEADER_BOARD_NAME & "] = '" & CB_Boards.Text & "' ORDER BY [" & DB_HEADER_REF_DES & "]"
		PCAD_BOM_da = New SqlDataAdapter(PCAD_BOM_myCmd)
		PCAD_BOM_ds = New DataSet()

		PCAD_BOM_da.Fill(PCAD_BOM_ds, 0, 500, "PCAD")

		DGV_PCAD_BOM.DataSource = Nothing
		DGV_PCAD_BOM.DataSource = PCAD_BOM_ds.Tables("PCAD")

		DGV_PCAD_BOM.Width = TP_BOM_compare.Width - 15
		DGV_PCAD_BOM.AutoResizeColumns(DataGridViewAutoSizeColumnMode.AllCells)

		DGV_QB_BOM.Visible = False

		L_Board.Text = "Database: " & CB_Boards.Text

		fromPCADdatabase = True
		fromSearch = False
		fromCompareItems = False
		fromCompareBOM = False
		B_CompareQBItems.Enabled = True
		B_CompareQBBOM.Enabled = True
		B_Compare_ALPHA.Enabled = True
	End Sub

	Private Sub B_SearchBOM_Click() Handles B_SearchBOM.Click
		Cursor = Cursors.WaitCursor
		If LOGDATA = True Then
			Try
				SearchBOM()
			Catch ex As Exception
				UnhandledExceptionMessage(ex)
			End Try
		Else
			SearchBOM()
		End If
		Cursor = Cursors.Default
	End Sub

	Private Sub SearchBOM()
		Dim selectFile As New OpenFileDialog()
		selectFile.InitialDirectory = My.Settings.BOMFilePath
		selectFile.Filter = "All files (*.*)|*.*|All files (*.*)|*.*"
		If selectFile.ShowDialog() = DialogResult.OK Then
			My.Settings.BOMFilePath = selectFile.FileName
			My.Settings.Save()

			Dim originalName As String = Path.GetFileName(selectFile.FileName)
			Dim fileNameParsed() As String = originalName.Split(".")

			If fileNameParsed.Length < 4 Then
				MsgBox("The file you have selected does not parse into 5 parts (boardname).(Rev#).(#).(Option).(bom)" & vbNewLine & "Please check the name of your file for correct naming conventions.")
				Return
			End If

			Dim fileName As String = fileNameParsed(INDEX_BOARD) & "." & fileNameParsed(INDEX_REVISION1) & "." & fileNameParsed(INDEX_REVISION2) & "." & fileNameParsed(INDEX_OPTION) & "."

			ParseBOMFile(My.Settings.BOMFilePath, False)

			PopulateDataTable(fileName)

			TB_FilePath.Text = My.Settings.BOMFilePath

			L_Board.Text = "File: " & fileName
		End If

		DGV_PCAD_BOM.Width = TP_BOM_compare.Width - 15
		DGV_PCAD_BOM.AutoResizeColumns(DataGridViewAutoSizeColumnMode.AllCells)

		DGV_QB_BOM.Visible = False

		fromPCADdatabase = False
		fromSearch = True
		fromCompareItems = False
		fromCompareBOM = False
		B_CompareQBItems.Enabled = True
		B_CompareQBBOM.Enabled = True
		B_Compare_ALPHA.Enabled = True
	End Sub

	Private Function ParseBOMFile(ByRef filePath As String, ByRef isRelease As Boolean) As Boolean
		TB_SearchIndicator.BackColor = Color.Black
		Dim errors As Boolean = False
		myCmd.CommandText = "DELETE FROM " & TABLE_TEMP_PCADBOM
		myCmd.ExecuteNonQuery()

		Dim originalName As String = Path.GetFileName(filePath)
		Dim fileNameParsed() As String = originalName.Split(".")
		If fileNameParsed.Length < 4 Then
			MsgBox("The file you have selected does not parse into 5 parts (boardname).(Rev#).(#).(Option).(bom)" & vbNewLine & "Please check the name of your file for correct naming conventions.")
			Return False
		End If
		Dim fileName As String = fileNameParsed(INDEX_BOARD) & "." & fileNameParsed(INDEX_REVISION1) & "." & fileNameParsed(INDEX_REVISION2) & "." & fileNameParsed(INDEX_OPTION) & "."

		'Indexs
		Dim INDEX_refdes As Integer = -1
		Dim INDEX_value As Integer = -1
		Dim INDEX_vendor As Integer = -1
		Dim INDEX_partNumber As Integer = -1
		Dim INDEX_stockNumber As Integer = -1
		Dim INDEX_process As Integer = -1

		'Optional
		Dim INDEX_options As Integer = -1
		Dim INDEX_swap As Integer = -1

		Dim swapList As New List(Of String)
		Dim optionList As New List(Of String)
		Dim hasPCB As Boolean = False

		'Start our transaction. Must assign both transaction object and connection to the command object for a pending local transaction.
		Dim transaction As SqlTransaction = Nothing
		transaction = myConn.BeginTransaction("Temp Transaction")
		myCmd.Connection = myConn
		myCmd.Transaction = transaction

		Try
			Using myParser As New TextFieldParser(filePath)
				myParser.TextFieldType = FieldType.Delimited
				myParser.SetDelimiters(",")
				Dim currentRow As String()

				'First row is the header row.
				currentRow = myParser.ReadFields()
				Dim index As Integer = 0

				'Parse the header row to grab Indexs. They can be generated in any order.
				For Each header In currentRow
					Select Case header.ToLower
						Case "refdes"
							INDEX_refdes = index
						Case "value"
							INDEX_value = index
						Case "vendor"
							INDEX_vendor = index
						Case "part number"
							INDEX_partNumber = index
						Case "stock number"
							INDEX_stockNumber = index
						Case "process"
							INDEX_process = index
						Case "option"
							INDEX_options = index
						Case "swap"
							INDEX_swap = index
					End Select
					index += 1
				Next

				Dim errorsPresent As String = ""
				Dim hasErrors As Boolean = False

				'Check to make sure that we have all of the headers that we need.
				If INDEX_refdes = -1 Then
					errorsPresent = """RefDes"""
					hasErrors = True
				End If
				If INDEX_value = -1 Then
					errorsPresent = errorsPresent & " ""Value"""
					hasErrors = True
				End If
				If INDEX_vendor = -1 Then
					errorsPresent = errorsPresent & " ""Vendor"""
					hasErrors = True
				End If
				If INDEX_partNumber = -1 Then
					errorsPresent = errorsPresent & " ""PartNumber"""
					hasErrors = True
				End If
				If INDEX_stockNumber = -1 Then
					errorsPresent = errorsPresent & " ""StockNumber"""
					hasErrors = True
				End If
				If INDEX_process = -1 Then
					errorsPresent = errorsPresent & " ""Process"""
					hasErrors = True
				End If

				If hasErrors = True Then
					TB_SearchIndicator.BackColor = Color.Red
					myCmd.CommandText = "INSERT INTO " & TABLE_TEMP_PCADBOM & " ([" & DB_HEADER_REF_DES & "]) VALUES('" & errorsPresent & "')"
					myCmd.ExecuteNonQuery()
					transaction.Commit()
					Return True
				End If

				While Not myParser.EndOfData
					'Parse Values
					Dim name As String = ""
					Dim referenceDesignator As String = ""
					Dim value As String = ""
					Dim vendor As String = ""
					Dim partNumber As String = ""
					Dim stockNumber As String = ""
					Dim process As String = ""
					Dim prefix As String = ""
					errorsPresent = ""

					'Optional
					Dim optionValue As String = ""
					Dim swapValue As String = ""
					currentRow = myParser.ReadFields()

					'If the row is blank, move on to the next row.
					If currentRow(0).Length = 0 Then
						Continue While
					End If

					'If we are dealing with options, an 'X' denotes that this part is not used. Move on to the next part.
					If INDEX_options <> -1 And fileNameParsed(INDEX_OPTION).Length <> 0 Then
						If String.Compare(currentRow(INDEX_options), "X") = 0 Then
							Continue While
						End If
					End If

					If currentRow(INDEX_process).ToString.ToUpper = PROCESS_NOTUSED Then
						Continue While
					End If

					'- - - Parse Reference Designator - - -

					referenceDesignator = currentRow(INDEX_refdes)

					'- - - Parse Value - - -

					value = currentRow(INDEX_value)

					'Check to see if we have single quotes and replace them with double single qoutes.
					If value.Contains("'"c) = True Then
						value = value.Replace("'", "''")
					End If

					'Check to see if we have any curly braces. -denotes default value.
					If value.Length = 0 Or value.Contains("{"c) = True Or value.Contains("}"c) = True Then
						errorsPresent = errorsPresent & "|""Value-Syntax"""
						hasErrors = True
					End If

					'Check to see if we are dealing with a 'ZF'.
					If referenceDesignator.Contains(REFERENCE_DESIGNATOR_FIDUCIAL) = True Then
						If String.Compare(value, "NO SMT", True) <> 0 Then
							errorsPresent = errorsPresent & "|""Value-Fiducial Syntax"""
							hasErrors = True
						End If
					End If

					'- - - Parse Vendor - - -

					vendor = currentRow(INDEX_vendor)

					'Check to see if we have any curly braces. -denotes default value.
					If vendor.Length = 0 Or vendor.Contains("{"c) = True Or vendor.Contains("}"c) = True Then
						'Check to see if we are dealing with a 'ZD' or 'ZF'.
						If referenceDesignator.Contains(REFERENCE_DESIGNATOR_OPTION) = False And referenceDesignator.Contains(REFERENCE_DESIGNATOR_FIDUCIAL) = False Then
							errorsPresent = errorsPresent & "|""Vendor-Syntax"""
							hasErrors = True
						End If
					End If

					'- - - Parse Part Number - - -

					partNumber = currentRow(INDEX_partNumber)

					'Check to see if we have any curly braces. -denotes default value.
					If partNumber.Length = 0 Or partNumber.Contains("{"c) = True Or partNumber.Contains("}"c) = True Then
						'Check to see if we are dealing with a 'ZD' or 'ZF'.
						If referenceDesignator.Contains(REFERENCE_DESIGNATOR_OPTION) = False And referenceDesignator.Contains(REFERENCE_DESIGNATOR_FIDUCIAL) = False Then
							errorsPresent = errorsPresent & "|""PartNumber-Syntax"""
							hasErrors = True
						End If
					End If

					'- - - Parse Stock Number - - -
					'Check to see that we have a ':'
					'Check to see that we have a value
					'Check to see that we do not have any curly braces. - denotes default value.
					If currentRow(INDEX_stockNumber).Contains(":") = True And
					currentRow(INDEX_stockNumber).Length <> 0 And
					currentRow(INDEX_stockNumber).Contains("{"c) = False And
					currentRow(INDEX_stockNumber).Contains("}"c) = False Then

						stockNumber = currentRow(INDEX_stockNumber).Substring(currentRow(INDEX_stockNumber).IndexOf(":") + 1)
					Else
						'Check to see if we are dealing with a 'ZD'.
						If referenceDesignator.Contains(REFERENCE_DESIGNATOR_OPTION) = False And referenceDesignator.Contains(REFERENCE_DESIGNATOR_FIDUCIAL) = False Then
							errorsPresent = errorsPresent & "|""Stock Number-Syntax"""
							hasErrors = True
							stockNumber = currentRow(INDEX_stockNumber)
						End If
					End If

					'- - - Parse the Prefix - - -

					'Check to see if we have a colon in the Stock Number.
					If currentRow(INDEX_stockNumber).Contains(":") = True Then
						prefix = currentRow(INDEX_stockNumber).Substring(0, currentRow(INDEX_stockNumber).IndexOf(":"))

						'Check to see if we are dealing with a 'PCB' item.
						If String.Compare(prefix, PREFIX_PCB, True) = 0 Then
							hasPCB = True
							name = stockNumber.Substring(stockNumber.IndexOf("-") + 1)

							'Check to see if our file name has any options.
							If fileNameParsed(INDEX_OPTION).Length = 0 Then
								'Check to see if our 'PCB' Stock name matches our filename.
								If String.Compare(name, fileName, True) <> 0 Then
									errorsPresent = errorsPresent & "|""FileName and PCB Stock do not match"""
									hasErrors = True
								End If

								'Check to see if our 'PCB' Part name matches our filename.
								name = partNumber.Substring(partNumber.IndexOf("-") + 1)
								If String.Compare(name, fileName, True) <> 0 Then
									errorsPresent = errorsPresent & "|""FileName and PCB Part do not match"""
									hasErrors = True
								End If
							End If
						End If
					Else
						'Check to see if we are dealing with a 'ZD' or 'ZF'.
						If referenceDesignator.Contains(REFERENCE_DESIGNATOR_OPTION) = False And referenceDesignator.Contains(REFERENCE_DESIGNATOR_FIDUCIAL) = False Then
							errorsPresent = errorsPresent & "|""Stock Prefix-Syntax"""
							hasErrors = True
						End If
					End If

					'- - - Parse the Process - - -

					process = currentRow(INDEX_process)

					'Check to see if we have any curly braces. -denotes default value.
					If process.Length = 0 Or process.Contains("{"c) = True Or process.Contains("}"c) = True Then
						'Check to see if we are dealing with a 'ZD'.
						If referenceDesignator.Contains(REFERENCE_DESIGNATOR_OPTION) = False And referenceDesignator.Contains(REFERENCE_DESIGNATOR_FIDUCIAL) = False Then
							errorsPresent = errorsPresent & "|""Process-Syntax"""
							hasErrors = True
						End If
					End If

					'Check to see if the process is valid.
					If String.Compare(process, PROCESS_SMT, True) <> 0 And
						String.Compare(process, PROCESS_SMTBOTTOM, True) <> 0 And
						String.Compare(process, PROCESS_HANDFLOW, True) <> 0 And
						String.Compare(process, PROCESS_POSTASSEMBLY, True) <> 0 And
						String.Compare(process, PROCESS_PCBBOARD, True) <> 0 And
						String.Compare(process, PROCESS_SMTHAND, True) <> 0 And
						String.Compare(process, PREFIX_BAS, True) <> 0 And
						String.Compare(process, PREFIX_BIS, True) <> 0 And
						String.Compare(process, PREFIX_DAS, True) <> 0 Then
						'Check to see if we are using a 'Not Used'.
						If String.Compare(process, PROCESS_NOTUSED, True) = 0 Then
							Continue While
						Else
							'Check to see if we are dealing with a 'ZD' or 'ZF'.
							If referenceDesignator.Contains(REFERENCE_DESIGNATOR_OPTION) = False And referenceDesignator.Contains(REFERENCE_DESIGNATOR_FIDUCIAL) = False Then
								errorsPresent = errorsPresent & "|""Process-Name"""
								hasErrors = True
							End If
						End If
					End If

					'Check to see if we have the option field or not.
					If INDEX_options <> -1 Then
						Dim include As Boolean = False

						'Get the files' options. 
						' Format is in single letters [ABD]
						' This means to include anything with an A|B|D inside its option feild.
						optionValue = currentRow(INDEX_options)

						'Check to see if we have a value first. If we do not have a value, then this part is needed accros all options.
						If optionValue.Length <> 0 Then

							'Check to see if we have this option as part of our list of valid options.
							If optionList.Contains(optionValue) = False Then
								optionList.Add(optionValue)
							End If

							'Next, check to see if our prefix is part of the Option Description
							If fileNameParsed(INDEX_OPTION).Length <> 0 Then
								For index = 0 To optionValue.Length - 1
									'Check each letter of the option feild to see if the file calls for it.
									If fileNameParsed(INDEX_OPTION).Contains(optionValue(index)) = True Then
										include = True
										Exit For
									End If
								Next

								If isRelease = False Then
									If include = False Then
										Continue While
									End If
								End If


								'Check to see if we are dealing with a 'ZD'.
								If referenceDesignator.Contains(REFERENCE_DESIGNATOR_OPTION) = True Then
									'Check to make sure the option is not an 'X'.
									If String.Compare(optionValue, "X") <> 0 Then
										'Check to see if the values match each other.
										If value.Substring(0, value.IndexOf("-")).Contains(optionValue) = False Then
											errorsPresent = errorsPresent & "|""Option-Syntax"""
											hasErrors = True
										End If
									End If
								End If
							End If
						End If
					End If

					'Check to see if we have the swap field or not.
					If INDEX_swap <> -1 Then
						swapValue = currentRow(INDEX_swap)

						'Check to see if we have a swap.
						If currentRow(INDEX_swap).Length <> 0 Then
							'Check to see if we are dealing with a 'ZD'.
							If referenceDesignator.Contains(REFERENCE_DESIGNATOR_OPTION) = True Or referenceDesignator.Contains(REFERENCE_DESIGNATOR_FIDUCIAL) = True Then
								errorsPresent = errorsPresent & "|""Option-Syntax"""
								hasErrors = True
							Else
								swapList.Add(currentRow(INDEX_swap))
							End If
						End If
					End If

					If hasErrors = True Then
						errors = True
					End If

					myCmd.CommandText = "INSERT INTO " & TABLE_TEMP_PCADBOM & " ([" & DB_HEADER_REF_DES & "], [" & DB_HEADER_DESCRIPTION & "], [" & DB_HEADER_VENDOR & "], [" & DB_HEADER_MPN & "], [" & DB_HEADER_ITEM_NUMBER & "], [" & DB_HEADER_PROCESS & "], [" & DB_HEADER_ITEM_PREFIX & "], [" & DB_HEADER_BOARD_NAME & "], [" & DB_HEADER_OPTION & "], [" & DB_HEADER_SWAP & "], [" & DB_HEADER_ERRORS & "]) " &
										"VALUES('" & referenceDesignator & "', '" & value & "', '" & vendor & "', '" & partNumber & "', '" & stockNumber & "', '" & process & "', '" & prefix & "', '" & fileName & "', '" & optionValue & "', '" & swapValue & "', '" & errorsPresent & "')"
					myCmd.ExecuteNonQuery()
				End While
			End Using

			'Check to see if we have a swap list.
			If swapList.Count <> 0 And fileNameParsed(INDEX_OPTION).Length = 0 Then
				'Check to see that each swap has a reference designator to swap with.
				For Each item In swapList
					myCmd.CommandText = "IF EXISTS(SELECT * FROM " & TABLE_TEMP_PCADBOM & " WHERE [" & DB_HEADER_REF_DES & "] = '" & item & "') SELECT 1 ELSE SELECT 0"
					Dim myReader As SqlDataReader = myCmd.ExecuteReader
					If myReader.Read() Then
						If myReader.GetInt32(0) = 0 Then
							myReader.Close()
							myCmd.CommandText = "INSERT INTO " & TABLE_TEMP_PCADBOM & " ([" & DB_HEADER_ERRORS & "]) " &
										"VALUES('""Swap " & item & """')"
							myCmd.ExecuteNonQuery()
						End If
					End If
					myReader.Close()
				Next
			End If

			'Grab all of the 'ZD' Values.
			myCmd.CommandText = "SELECT [" & DB_HEADER_DESCRIPTION & "] FROM " & TABLE_TEMP_PCADBOM & " WHERE [" & DB_HEADER_REF_DES & "] LIKE '" & REFERENCE_DESIGNATOR_OPTION & "%'"
			Dim bomcheck_da = New SqlDataAdapter(myCmd)
			Dim bomcheck_ds = New DataSet()
			bomcheck_da.Fill(bomcheck_ds, "bom DATA")

			'Check to see if we have any option.
			If optionList.Count <> 0 And fileNameParsed(INDEX_OPTION).Length = 0 Then
				For Each item In optionList
					'Check to see if our option contains an 'X'. If yes then continue to the next item.
					If item.Contains("X") = True Then
						Continue For
					End If
					Dim found As Boolean = False

					'Check to see if we have a 'ZD' for each option.
					For Each dsRow As DataRow In bomcheck_ds.Tables("bom DATA").Rows
						Dim optionDescription As String = dsRow(DB_HEADER_DESCRIPTION).ToString.Substring(0, dsRow(DB_HEADER_DESCRIPTION).ToString.IndexOf("-")).Trim

						For index = 0 To item.Length - 1
							'Check each letter of the option feild to see if the file calls for it.
							For index2 = 0 To optionDescription.Length - 1
								If optionDescription(index2).Equals(item(index)) = True Then
									found = True
									Exit For
								End If
							Next
							If found = True Then
								Exit For
							End If
						Next

						If found = True Then
							Exit For
						End If
					Next

					If found = False Then
						myCmd.CommandText = "INSERT INTO " & TABLE_TEMP_PCADBOM & " ([" & DB_HEADER_ERRORS & "]) " &
										"VALUES('""Option " & item & " Description""')"
						myCmd.ExecuteNonQuery()
					End If
				Next
			End If

			'Check to see if we have a 'PCB' item.
			If hasPCB = False Then
				myCmd.CommandText = "INSERT INTO " & TABLE_TEMP_PCADBOM & " ([" & DB_HEADER_ERRORS & "]) " &
										"VALUES('""PCB Designator""')"
				myCmd.ExecuteNonQuery()
			End If
			transaction.Commit()

			'Check to see if we had any errors.
			If errors = False Then
				TB_SearchIndicator.BackColor = Color.LightGreen
			Else
				TB_SearchIndicator.BackColor = Color.Red
			End If
			Return True
		Catch ex As Exception
			If Not transaction Is Nothing Then
				sqlapi.RollBack(transaction, errorMessage:=New List(Of String))
				MsgBox(ex.Message)
				Return False
			End If
		End Try
		Return True
	End Function

	Private Sub PopulateDataTable(ByRef filePath As String)
		PCAD_BOM_myCmd.CommandText = "SELECT [" & DB_HEADER_REF_DES & "], [" & DB_HEADER_DESCRIPTION & "], [" & DB_HEADER_MPN & "], COALESCE(NULLIF([" & DB_HEADER_ITEM_PREFIX & "] + ':' + [" & DB_HEADER_ITEM_NUMBER & "],':'), [" & DB_HEADER_ITEM_NUMBER & "]) AS '" & DB_HEADER_ITEM_NUMBER & "', [" & DB_HEADER_VENDOR & "], [" & DB_HEADER_PROCESS & "], [" & DB_HEADER_OPTION & "], [" & DB_HEADER_SWAP & "], [" & DB_HEADER_ERRORS & "] FROM " & TABLE_TEMP_PCADBOM & " ORDER BY [" & DB_HEADER_REF_DES & "]"
		PCAD_BOM_da = New SqlDataAdapter(PCAD_BOM_myCmd)
		PCAD_BOM_ds = New DataSet()

		PCAD_BOM_da.Fill(PCAD_BOM_ds, 0, 500, "PCAD")

		DGV_PCAD_BOM.DataSource = Nothing
		DGV_PCAD_BOM.DataSource = PCAD_BOM_ds.Tables("PCAD")
	End Sub

	Private Sub B_ReloadSearch_Click() Handles B_ReloadSearch.Click
		Cursor = Cursors.WaitCursor
		If LOGDATA = True Then
			Try
				ReloadSearch()
			Catch ex As Exception
				UnhandledExceptionMessage(ex)
			End Try
		Else
			ReloadSearch()
		End If
		Cursor = Cursors.Default
	End Sub

	Private Sub ReloadSearch()
		Dim originalName As String = Path.GetFileName(My.Settings.BOMFilePath.Substring(My.Settings.BOMFilePath.LastIndexOf("\")))
		Dim fileNameParsed() As String = originalName.Split(".")
		If fileNameParsed.Length < 4 Then
			MsgBox("The file you have selected does not parse into 5 parts (boardname).(Rev#).(#).(Option).(bom)" & vbNewLine & "Please check the name of your file for correct naming conventions.")
			Return
		End If
		Dim fileName As String = fileNameParsed(INDEX_BOARD) & "." & fileNameParsed(INDEX_REVISION1) & "." & fileNameParsed(INDEX_REVISION2) & "." & fileNameParsed(INDEX_OPTION) & "."

		L_Board.Text = "File: " & fileName

		myCmd.CommandText = "DELETE FROM " & TABLE_TEMP_PCADBOM
		myCmd.ExecuteNonQuery()

		ParseBOMFile(My.Settings.BOMFilePath, True)

		PopulateDataTable(fileName)

		TB_FilePath.Text = My.Settings.BOMFilePath

		DGV_PCAD_BOM.Width = TP_BOM_compare.Width - 15
		DGV_PCAD_BOM.AutoResizeColumns(DataGridViewAutoSizeColumnMode.AllCells)

		DGV_QB_BOM.Visible = False

		fromPCADdatabase = False
		fromSearch = True
		fromCompareItems = False
		fromCompareBOM = False
		B_CompareQBItems.Enabled = True
		B_CompareQBBOM.Enabled = True
		B_Compare_ALPHA.Enabled = True
	End Sub

	Private Sub B_CompareQBItems_Click() Handles B_CompareQBItems.Click
		Cursor = Cursors.WaitCursor
		If LOGDATA = True Then
			Try
				If ChangeCheck(True) = True Then
					CompareQBItems()
				End If
			Catch ex As Exception
				UnhandledExceptionMessage(ex)
			End Try
		Else
			If ChangeCheck(True) = True Then
				CompareQBItems()
			End If
		End If
		Cursor = Cursors.Default
	End Sub

	Private Sub CompareQBItems()
		Dim message As String = ""
		If sqlapi.CheckDirtyBit(message) = True Then
			MsgBox(message)
			Return
		End If

		TB_SearchIndicator.BackColor = Color.Black
		Dim DataTable_QB_Compare As New DataTable
		DataTable_QB_Compare.Columns.Add(DB_HEADER_ITEM_PREFIX, GetType(String))
		DataTable_QB_Compare.Columns.Add(DB_HEADER_ITEM_NUMBER, GetType(String))
		DataTable_QB_Compare.Columns.Add(DB_HEADER_MPN, GetType(String))
		DataTable_QB_Compare.Columns.Add(DB_HEADER_VENDOR, GetType(String))
		DataTable_QB_Compare.Columns.Add(DB_HEADER_MPN2, GetType(String))
		DataTable_QB_Compare.Columns.Add(DB_HEADER_VENDOR2, GetType(String))
		DataTable_QB_Compare.Columns.Add(DB_HEADER_MPN3, GetType(String))
		DataTable_QB_Compare.Columns.Add(DB_HEADER_VENDOR3, GetType(String))

		Dim DataTable_PCAD_Compare As New DataTable
		DataTable_PCAD_Compare.Columns.Add(DB_HEADER_REF_DES, GetType(String))
		DataTable_PCAD_Compare.Columns.Add(DB_HEADER_ITEM_PREFIX, GetType(String))
		DataTable_PCAD_Compare.Columns.Add(DB_HEADER_ITEM_NUMBER, GetType(String))
		DataTable_PCAD_Compare.Columns.Add(DB_HEADER_MPN, GetType(String))
		DataTable_PCAD_Compare.Columns.Add(DB_HEADER_VENDOR, GetType(String))

		Dim QBItemList As New DataTable
		myCmd.CommandText = "SELECT * FROM " & TABLE_QB_ITEMS
		QBItemList.Load(myCmd.ExecuteReader())

		'Figure out if we are comparing a database entry or a search.
		If fromPCADdatabase Then
			'Create and add our query into our data set for comparison.
			PCAD_BOM_myCmd.CommandText = "SELECT [" & DB_HEADER_REF_DES & "], [" & DB_HEADER_ITEM_PREFIX & "], [" & DB_HEADER_ITEM_NUMBER & "], [" & DB_HEADER_MPN & "], [" & DB_HEADER_VENDOR & "] FROM " & TABLE_PCADBOM & " WHERE [" & DB_HEADER_BOARD_NAME & "] = '" & CB_Boards.Text & "' ORDER BY [" & DB_HEADER_ITEM_NUMBER & "]"
		ElseIf fromSearch Then
			'Create and add our query into our data set for comparison.
			Dim fileInformation As New FileInfo(My.Settings.BOMFilePath)
			Dim fileNameParsed() As String = fileInformation.Name.Split(".")
			If fileNameParsed.Length < 4 Then
				MsgBox("The file you have selected does not parse into 5 parts (boardname).(Rev#).(#).(Option).(bom)" & vbNewLine & "Please check the name of your file for correct naming conventions.")
				Return
			End If
			Dim fileName As String = fileNameParsed(INDEX_BOARD) & "." & fileNameParsed(INDEX_REVISION1) & "." & fileNameParsed(INDEX_REVISION2) & "." & fileNameParsed(INDEX_OPTION) & "."
			PCAD_BOM_myCmd.CommandText = "SELECT [" & DB_HEADER_REF_DES & "], [" & DB_HEADER_ITEM_PREFIX & "], [" & DB_HEADER_ITEM_NUMBER & "], [" & DB_HEADER_MPN & "], [" & DB_HEADER_VENDOR & "] FROM " & TABLE_TEMP_PCADBOM & " WHERE [" & DB_HEADER_BOARD_NAME & "] = '" & fileName & "' ORDER BY [" & DB_HEADER_ITEM_NUMBER & "]"
		End If

		PCAD_BOM_da = New SqlDataAdapter(PCAD_BOM_myCmd)
		PCAD_BOM_ds = New DataSet()
		PCAD_BOM_da.Fill(PCAD_BOM_ds, "PCAD DATA")

		For Each dsRow As DataRow In PCAD_BOM_ds.Tables("PCAD DATA").Rows
			'Check to see if we are dealing with 'ZD' or 'ZF'.
			If dsRow(DB_HEADER_REF_DES).Contains(REFERENCE_DESIGNATOR_OPTION) = True Or dsRow(DB_HEADER_REF_DES).Contains(REFERENCE_DESIGNATOR_FIDUCIAL) = True Then
				Continue For
			End If

			If CkB_OnlyDifferences.Checked = False Then
				'Show everything 
				'Check to see if the Stock Number exists.

				Dim itemdrs() As DataRow = QBItemList.Select("[" & DB_HEADER_ITEM_NUMBER & "] = '" & dsRow(DB_HEADER_ITEM_NUMBER) & "' AND [" & DB_HEADER_ITEM_PREFIX & "] = '" & dsRow(DB_HEADER_ITEM_PREFIX) & "'")

				If itemdrs.Length <> 0 Then
					'The stock number is inside the database 
					DataTable_QB_Compare.Rows.Add(itemdrs(0)(DB_HEADER_ITEM_PREFIX), itemdrs(0)(DB_HEADER_ITEM_NUMBER), itemdrs(0)(DB_HEADER_MPN), itemdrs(0)(DB_HEADER_VENDOR), itemdrs(0)(DB_HEADER_MPN2), itemdrs(0)(DB_HEADER_VENDOR2), itemdrs(0)(DB_HEADER_MPN3), itemdrs(0)(DB_HEADER_VENDOR3))
					DataTable_PCAD_Compare.Rows.Add(dsRow(DB_HEADER_REF_DES), dsRow(DB_HEADER_ITEM_PREFIX), dsRow(DB_HEADER_ITEM_NUMBER), dsRow(DB_HEADER_MPN), dsRow(DB_HEADER_VENDOR))
				Else
					'The stock number is not inside the database..
					DataTable_PCAD_Compare.Rows.Add(dsRow(DB_HEADER_REF_DES), dsRow(DB_HEADER_ITEM_PREFIX), dsRow(DB_HEADER_ITEM_NUMBER), dsRow(DB_HEADER_MPN), dsRow(DB_HEADER_VENDOR))
				End If
			Else
				If CkB_Include.Checked = True Then
					'Show advance diff. Limited to stock number, prefix, manufacture 1, part number 1
					'Check to see if the Stock Number exists.
					Dim exists() As DataRow = QBItemList.Select("[" & DB_HEADER_ITEM_NUMBER & "] = '" & dsRow(DB_HEADER_ITEM_NUMBER) & "' AND [" & DB_HEADER_ITEM_PREFIX & "] = '" & dsRow(DB_HEADER_ITEM_PREFIX) & "' AND (" &
						"([" & DB_HEADER_VENDOR & "] = '" & dsRow(DB_HEADER_VENDOR) & "' AND [" & DB_HEADER_MPN & "] = '" & dsRow(DB_HEADER_MPN) & "') OR " &
						"([" & DB_HEADER_VENDOR2 & "] = '" & dsRow(DB_HEADER_VENDOR) & "' AND [" & DB_HEADER_MPN2 & "] = '" & dsRow(DB_HEADER_MPN) & "') OR " &
						"([" & DB_HEADER_VENDOR3 & "] = '" & dsRow(DB_HEADER_VENDOR) & "' AND [" & DB_HEADER_MPN3 & "] = '" & dsRow(DB_HEADER_MPN) & "')) ")

					Dim itemdrs() As DataRow = QBItemList.Select("[" & DB_HEADER_ITEM_NUMBER & "] = '" & dsRow(DB_HEADER_ITEM_NUMBER) & "' AND [" & DB_HEADER_ITEM_PREFIX & "] = '" & dsRow(DB_HEADER_ITEM_PREFIX) & "'")

					'Check to see if we do not have a one-to-one match.
					If exists.Length = 0 Then
						'Check to see if the item exists in the database.
						If itemdrs.Length <> 0 Then
							'The stock number exists in the database but does not match the manufacture or part number.
							DataTable_PCAD_Compare.Rows.Add(dsRow(DB_HEADER_REF_DES), dsRow(DB_HEADER_ITEM_PREFIX), dsRow(DB_HEADER_ITEM_NUMBER), dsRow(DB_HEADER_MPN), dsRow(DB_HEADER_VENDOR))
							DataTable_QB_Compare.Rows.Add(itemdrs(0)(DB_HEADER_ITEM_PREFIX), itemdrs(0)(DB_HEADER_ITEM_NUMBER), itemdrs(0)(DB_HEADER_MPN), itemdrs(0)(DB_HEADER_VENDOR), itemdrs(0)(DB_HEADER_MPN2), itemdrs(0)(DB_HEADER_VENDOR2), itemdrs(0)(DB_HEADER_MPN3), itemdrs(0)(DB_HEADER_VENDOR3))
						Else
							'The stock number is not in the database.
							DataTable_PCAD_Compare.Rows.Add(dsRow(DB_HEADER_REF_DES), dsRow(DB_HEADER_ITEM_PREFIX), dsRow(DB_HEADER_ITEM_NUMBER), dsRow(DB_HEADER_MPN), dsRow(DB_HEADER_VENDOR))
						End If
					End If
				Else
					'Show a basic diff. Only limited to Stock number and prefix.
					'Check to see if the Stock Number exists.
					Dim itemdrs() As DataRow = QBItemList.Select("[" & DB_HEADER_ITEM_NUMBER & "] = '" & dsRow(DB_HEADER_ITEM_NUMBER) & "' AND [" & DB_HEADER_ITEM_PREFIX & "] = '" & dsRow(DB_HEADER_ITEM_PREFIX) & "'")

					If itemdrs.Length = 0 Then
						'The stock number is not inside the database..
						DataTable_PCAD_Compare.Rows.Add(dsRow(DB_HEADER_REF_DES), dsRow(DB_HEADER_ITEM_PREFIX), dsRow(DB_HEADER_ITEM_NUMBER), dsRow(DB_HEADER_MPN), dsRow(DB_HEADER_VENDOR))
					End If
				End If
			End If
		Next

		'Set our flag to show that we did a compare Items report for the excel button.
		fromCompareBOM = False
		fromCompareItems = True

		'Check to see if we added any rows. If we did not then there are no differences.
		If DataTable_PCAD_Compare.Rows.Count = 0 And DataTable_QB_Compare.Rows.Count = 0 Then
			DataTable_PCAD_Compare.Rows.Add("There are no differences between the PCAD BOM and the Items in QB.")
			DataTable_QB_Compare.Rows.Add("There are no differences between the PCAD BOM and the Items in QB.")
		End If

		PCAD_BOM_ds = New DataSet
		PCAD_BOM_ds.Tables.Add(DataTable_PCAD_Compare)
		DGV_PCAD_BOM.DataSource = Nothing
		DGV_PCAD_BOM.DataSource = DataTable_PCAD_Compare
		DGV_PCAD_BOM.AutoResizeColumns(DataGridViewAutoSizeColumnMode.AllCells)

		QB_BOM_ds = New DataSet
		QB_BOM_ds.Tables.Add(DataTable_QB_Compare)
		DGV_QB_BOM.DataSource = Nothing
		DGV_QB_BOM.DataSource = DataTable_QB_Compare
		DGV_QB_BOM.Visible = True
		DGV_QB_BOM.AutoResizeColumns(DataGridViewAutoSizeColumnMode.AllCells)

		ResizeTables(TP_BOM_compare)
	End Sub

	Private Sub B_CompareQBBOM_Click() Handles B_CompareQBBOM.Click
		Cursor = Cursors.WaitCursor
		If LOGDATA = True Then
			Try
				If ChangeCheck(True) = True Then
					CompareQBBOM()
				End If
			Catch ex As Exception
				UnhandledExceptionMessage(ex)
			End Try
		Else
			If ChangeCheck(True) = True Then
				CompareQBBOM()
			End If
		End If
		Cursor = Cursors.Default
	End Sub

	Private Sub CompareQBBOM()
		Dim message As String = ""
		If sqlapi.CheckDirtyBit(message) = True Then
			MsgBox(message)
			Return
		End If

		TB_SearchIndicator.BackColor = Color.Black

		'Set up the tables that we are going to be using.
		Dim DataTable_PCAD_Compare As New DataTable
		DataTable_PCAD_Compare.Columns.Add(DB_HEADER_REF_DES, GetType(String))
		DataTable_PCAD_Compare.Columns.Add(DB_HEADER_ITEM_PREFIX, GetType(String))
		DataTable_PCAD_Compare.Columns.Add(DB_HEADER_ITEM_NUMBER, GetType(String))
		DataTable_PCAD_Compare.Columns.Add(DB_HEADER_DESCRIPTION, GetType(String))
		DataTable_PCAD_Compare.Columns.Add(DB_HEADER_VENDOR, GetType(String))
		DataTable_PCAD_Compare.Columns.Add(DB_HEADER_MPN, GetType(String))
		DataTable_PCAD_Compare.Columns.Add(DB_HEADER_PROCESS, GetType(String))

		Dim DataTable_QB_Compare As New DataTable
		DataTable_QB_Compare.Columns.Add(DB_HEADER_ITEM_PREFIX, GetType(String))
		DataTable_QB_Compare.Columns.Add(DB_HEADER_ITEM_NUMBER, GetType(String))
		DataTable_QB_Compare.Columns.Add(DB_HEADER_DESCRIPTION, GetType(String))
		DataTable_QB_Compare.Columns.Add(DB_HEADER_VENDOR, GetType(String))
		DataTable_QB_Compare.Columns.Add(DB_HEADER_MPN, GetType(String))
		DataTable_QB_Compare.Columns.Add(DB_HEADER_PROCESS, GetType(String))

		Dim DataTable_PCAD_Quantity As New DataTable
		DataTable_PCAD_Quantity.Columns.Add(DB_HEADER_ITEM_NUMBER, GetType(String))
		DataTable_PCAD_Quantity.Columns.Add(HEADER_QTY_PCAD, GetType(String))
		DataTable_PCAD_Quantity.Columns.Add(HEADER_QTY_QB, GetType(String))

		Dim DataTable_QB_Quantity As New DataTable
		DataTable_QB_Quantity.Columns.Add(DB_HEADER_ITEM_NUMBER, GetType(String))
		DataTable_QB_Quantity.Columns.Add(HEADER_QTY_QB, GetType(String))
		DataTable_QB_Quantity.Columns.Add(HEADER_QTY_PCAD, GetType(String))

		Dim DataTable_PCAD_items_missing As New DataTable
		DataTable_PCAD_items_missing.Columns.Add(DB_HEADER_ITEM_NUMBER)

		Dim board As String = ""
		Dim table As String = ""
		Dim optionValue As String = ""

		'Figure out if we are comparing a database entry or a search.
		If fromPCADdatabase Then
			'Create and add our query into our data set for comparison.
			PCAD_BOM_myCmd.CommandText = "SELECT [" & DB_HEADER_REF_DES & "], [" & DB_HEADER_ITEM_PREFIX & "], [" & DB_HEADER_ITEM_NUMBER & "], [" & DB_HEADER_MPN & "], [" & DB_HEADER_VENDOR & "], [" & DB_HEADER_DESCRIPTION & "], [" & DB_HEADER_PROCESS & "] FROM " & TABLE_PCADBOM & " WHERE [" & DB_HEADER_BOARD_NAME & "] = '" & CB_Boards.Text & "' AND [" & DB_HEADER_PROCESS & "] != '" & PROCESS_NOTUSED & "' AND [" & DB_HEADER_PROCESS & "] != '" & PROCESS_POSTASSEMBLY & "' ORDER BY [" & DB_HEADER_ITEM_NUMBER & "]"
			board = CB_Boards.Text
			table = TABLE_PCADBOM
		ElseIf fromSearch Then
			'Create and add our query into our data set for comparison.
			Dim fileInformation As New FileInfo(My.Settings.BOMFilePath)
			Dim fileNameParsed() As String = fileInformation.Name.Split(".")
			If fileNameParsed.Length < 4 Then
				MsgBox("The file you have selected does not parse into 5 parts (boardname).(Rev#).(#).(Option).(bom)" & vbNewLine & "Please check the name of your file for correct naming conventions.")
				Return
			End If
			Dim fileName As String = fileNameParsed(INDEX_BOARD) & "." & fileNameParsed(INDEX_REVISION1) & "." & fileNameParsed(INDEX_REVISION2) & "." & fileNameParsed(INDEX_OPTION) & "."
			optionValue = fileNameParsed(INDEX_OPTION)
			PCAD_BOM_myCmd.CommandText = "SELECT [" & DB_HEADER_REF_DES & "], [" & DB_HEADER_ITEM_PREFIX & "], [" & DB_HEADER_ITEM_NUMBER & "], [" & DB_HEADER_MPN & "], [" & DB_HEADER_VENDOR & "], [" & DB_HEADER_DESCRIPTION & "], [" & DB_HEADER_PROCESS & "], [" & DB_HEADER_OPTION & "] FROM " & TABLE_TEMP_PCADBOM & " WHERE [" & DB_HEADER_BOARD_NAME & "] = '" & fileName & "' AND [" & DB_HEADER_PROCESS & "] != '" & PROCESS_NOTUSED & "' AND [" & DB_HEADER_OPTION & "] != 'X' AND [" & DB_HEADER_PROCESS & "] != '" & PROCESS_POSTASSEMBLY & "' ORDER BY [" & DB_HEADER_ITEM_NUMBER & "]"
			board = Path.GetFileName(fileName)
			table = TABLE_TEMP_PCADBOM
		End If

		PCAD_BOM_da = New SqlDataAdapter(PCAD_BOM_myCmd)
		PCAD_BOM_ds = New DataSet()
		PCAD_BOM_da.Fill(PCAD_BOM_ds, "PCAD DATA")

		Dim QBItemList As New DataTable
		myCmd.CommandText = "SELECT * FROM " & TABLE_QB_ITEMS
		QBItemList.Load(myCmd.ExecuteReader())

		QB_BOM_myCmd.commandText = "SELECT * FROM " & TABLE_QBBOM & " WHERE [" & DB_HEADER_NAME & "] = '" & board & "' ORDER BY [" & DB_HEADER_ITEM_NUMBER & "]"
		QB_BOM_da = New SqlDataAdapter(QB_BOM_myCmd)
		QB_BOM_ds = New DataSet()
		QB_BOM_da.Fill(QB_BOM_ds, "QB BOM DATA")

		If QB_BOM_ds.Tables("QB BOM DATA").Rows.Count <> 0 Then
			'---------------------------------'
			'Compare QB with our PCAD record. '
			'---------------------------------'

			For Each dsRow As DataRow In PCAD_BOM_ds.Tables("PCAD DATA").Rows
				Dim tempProcess As String = ""
				'Check to see if we are dealing with 'ZD'.
				If dsRow(DB_HEADER_REF_DES).Contains(REFERENCE_DESIGNATOR_OPTION) = True Or dsRow(DB_HEADER_REF_DES).Contains(REFERENCE_DESIGNATOR_FIDUCIAL) = True Then
					Continue For
				End If

				Dim optionColumn As String = ""

				If fromSearch = True Then
					optionColumn = dsRow(DB_HEADER_OPTION)
				End If

				Dim include As Boolean = True
				'Check to see if we have an option field to account for from our file search.
				If optionColumn.Length <> 0 And fromSearch = True Then
					For index = 0 To optionColumn.Length - 1
						For index2 = 0 To optionValue.Length - 1
							If optionValue(index2) = (optionColumn(index)) Then
								'We have been found and need to look no farther.
								include = True
								Exit For
							Else
								'Keep setting to false. If found we set to true and then exit.
								include = False
							End If
						Next
						If include = True Then
							Exit For
						End If
					Next
				End If

				If include = False Then
					Continue For
				End If

				'Check to see what process we are using. Will change what we grab from PCAD.
				If dsRow(DB_HEADER_PROCESS).ToString.Contains(PROCESS_SMT) = True Then
					tempProcess = "[" & DB_HEADER_PROCESS & "] LIKE '%" & PROCESS_SMT & "'"
				Else
					tempProcess = "[" & DB_HEADER_PROCESS & "] = '" & dsRow(DB_HEADER_PROCESS) & "'"
				End If

				Dim QBBOMfound() As DataRow = QB_BOM_ds.Tables("QB BOM DATA").Select("[" & DB_HEADER_ITEM_NUMBER & "] = '" & dsRow(DB_HEADER_ITEM_NUMBER) & "' AND [" & DB_HEADER_ITEM_PREFIX & "] = '" & dsRow(DB_HEADER_ITEM_PREFIX) & "' AND " & tempProcess)
				Dim QBquantity As Integer = 0

				If QBBOMfound.Length = 0 Then
					'The item is not in the QB BOM.
					DataTable_PCAD_Compare.Rows.Add(dsRow(DB_HEADER_REF_DES), dsRow(DB_HEADER_ITEM_PREFIX), dsRow(DB_HEADER_ITEM_PREFIX) & ":" & dsRow(DB_HEADER_ITEM_NUMBER), dsRow(DB_HEADER_DESCRIPTION), dsRow(DB_HEADER_VENDOR), dsRow(DB_HEADER_MPN), dsRow(DB_HEADER_PROCESS))
				Else
					QBquantity = QBBOMfound(0)(DB_HEADER_QUANTITY)
				End If

				'Check our quantity differences.

				'Dim PCADquantity As Integer = PCAD_BOM_ds.Tables("PCAD DATA").Select("[" & DB_HEADER_ITEM_NUMBER & "] = '" & dsRow(DB_HEADER_ITEM_NUMBER) & "'")

				Dim PCADquantity As Integer = 0
				For Each drow As DataRow In PCAD_BOM_ds.Tables("PCAD DATA").Select("[" & DB_HEADER_ITEM_NUMBER & "] = '" & dsRow(DB_HEADER_ITEM_NUMBER) & "'")
					include = True

					'Check to see if we have an option field to account for from our file search.
					If optionColumn.Length <> 0 And fromSearch = True Then
						For index = 0 To optionColumn.Length - 1
							For index2 = 0 To optionValue.Length - 1
								If optionValue(index2) = (optionColumn(index)) Then
									'We have been found and need to look no farther.
									include = True
									Exit For
								Else
									'Keep setting to false. If found we set to true and then exit.
									include = False
								End If
							Next
							If include = True Then
								Exit For
							End If
						Next
					End If

					If include = True Then
						PCADquantity += 1
					End If
				Next

				If PCADquantity <> QBquantity Then
					If DataTable_PCAD_Quantity.Select("[" & DB_HEADER_ITEM_NUMBER & "] = '" & dsRow(DB_HEADER_ITEM_NUMBER) & "'").Length = 0 Then
						DataTable_PCAD_Quantity.Rows.Add(dsRow(DB_HEADER_ITEM_NUMBER), PCADquantity, QBquantity)
					End If
				End If

				'Check to see if the Item is found in the QB Items list.
				If QBItemList.Select("[" & DB_HEADER_ITEM_NUMBER & "] = '" & dsRow(DB_HEADER_ITEM_NUMBER) & "' AND [" & DB_HEADER_ITEM_PREFIX & "] = '" & dsRow(DB_HEADER_ITEM_PREFIX) & "'").Length = 0 Then
					If DataTable_PCAD_items_missing.Select("[" & DB_HEADER_ITEM_NUMBER & "] = '" & dsRow(DB_HEADER_ITEM_PREFIX) & ":" & dsRow(DB_HEADER_ITEM_NUMBER) & "'").Length = 0 Then
						DataTable_PCAD_items_missing.Rows.Add(dsRow(DB_HEADER_ITEM_PREFIX) & ":" & dsRow(DB_HEADER_ITEM_NUMBER))
					End If
				End If
			Next

			'---------------------------------'
			'Compare PCAD with our QB record. '
			'---------------------------------'
			For Each dsRow As DataRow In QB_BOM_ds.Tables("QB BOM DATA").Rows
				'Check to see if we are dealing with an assembly item. If we are then we need to ignore it.
				If dsRow(DB_HEADER_ITEM_PREFIX).ToString.Contains(PREFIX_BAS) = True Or dsRow(DB_HEADER_ITEM_PREFIX).ToString.Contains(PREFIX_BIS) = True Or dsRow(DB_HEADER_ITEM_PREFIX).ToString.Contains(PREFIX_DAS) = True Or dsRow(DB_HEADER_ITEM_PREFIX).ToString.Contains(PREFIX_SMA) = True Then
					Continue For
				End If

				'Build our query string according to what database we are accessing.
				Dim query As String = "[" & DB_HEADER_ITEM_NUMBER & "] = '" & dsRow(DB_HEADER_ITEM_NUMBER) & "' AND [" & DB_HEADER_ITEM_PREFIX & "] = '" & dsRow(DB_HEADER_ITEM_PREFIX) & "' AND [" & DB_HEADER_PROCESS & "] LIKE '%" & dsRow(DB_HEADER_PROCESS) & "%'"

				Dim drs() As DataRow = PCAD_BOM_ds.Tables("PCAD DATA").Select(query)
				Dim itemdrs() As DataRow = QBItemList.Select("[" & DB_HEADER_ITEM_NUMBER & "] = '" & dsRow(DB_HEADER_ITEM_NUMBER) & "'")

				Dim PCADquantity As Integer = drs.Length
				Dim QBquantity As Integer = dsRow(DB_HEADER_QUANTITY)

				If PCADquantity = 0 Then
					'This item is not found in the PCAD Database. Check to see if it is in the Items database.

					If itemdrs.Length <> 0 Then
						'The item is in the database.
						DataTable_QB_Compare.Rows.Add(dsRow(DB_HEADER_ITEM_PREFIX), dsRow(DB_HEADER_ITEM_PREFIX) & ":" & dsRow(DB_HEADER_ITEM_NUMBER), itemdrs(0)(DB_HEADER_DESCRIPTION), itemdrs(0)(DB_HEADER_VENDOR), itemdrs(0)(DB_HEADER_MPN), dsRow(DB_HEADER_PROCESS))
					Else
						'The item is not in the database.
						DataTable_QB_Compare.Rows.Add(dsRow(DB_HEADER_ITEM_PREFIX), dsRow(DB_HEADER_ITEM_PREFIX) & ":" & dsRow(DB_HEADER_ITEM_NUMBER), "", "", "")
					End If
				End If

				If PCADquantity <> QBquantity Then
					If DataTable_QB_Quantity.Select("[" & DB_HEADER_ITEM_NUMBER & "] = '" & dsRow(DB_HEADER_ITEM_NUMBER) & "'").Length = 0 Then
						DataTable_QB_Quantity.Rows.Add(dsRow(DB_HEADER_ITEM_NUMBER), QBquantity, PCADquantity)
					End If
				End If
			Next
		Else
			DataTable_PCAD_Compare.Rows.Add("There is no QB BOM for: " & board)
			DataTable_QB_Compare.Rows.Add("There is no QB BOM for: " & board)
		End If

		If DataTable_PCAD_Quantity.Rows.Count <> 0 Or DataTable_QB_Quantity.Rows.Count <> 0 Then
			DataTable_PCAD_Compare.Rows.Add()
			DataTable_PCAD_Compare.Rows.Add("", "", DB_HEADER_ITEM_NUMBER, HEADER_QTY_PCAD, HEADER_QTY_QB)

			DataTable_QB_Compare.Rows.Add()
			DataTable_QB_Compare.Rows.Add("", "", DB_HEADER_ITEM_NUMBER, HEADER_QTY_QB, HEADER_QTY_PCAD)

			For Each row In DataTable_PCAD_Quantity.Rows
				DataTable_PCAD_Compare.Rows.Add("", "", row(DB_HEADER_ITEM_NUMBER), row(HEADER_QTY_PCAD), row(HEADER_QTY_QB))
			Next

			For Each row In DataTable_QB_Quantity.Rows
				DataTable_QB_Compare.Rows.Add("", "", row(DB_HEADER_ITEM_NUMBER), row(HEADER_QTY_QB), row(HEADER_QTY_PCAD))
			Next
		End If

		If DataTable_PCAD_items_missing.Rows.Count <> 0 Then
			DataTable_PCAD_Compare.Rows.Add()
			DataTable_PCAD_Compare.Rows.Add("Item Number Not in QB Items")

			For Each row In DataTable_PCAD_items_missing.Rows
				DataTable_PCAD_Compare.Rows.Add(row(DB_HEADER_ITEM_NUMBER))
			Next
		End If

		fromCompareBOM = True
		fromCompareItems = False

		If DataTable_PCAD_Compare.Rows.Count = 0 And DataTable_QB_Compare.Rows.Count = 0 Then
			DataTable_PCAD_Compare.Rows.Add("There are no missing components in the QB BOM.")
			DataTable_QB_Compare.Rows.Add("There are no extra components in the QB BOM.")
		End If

		PCAD_BOM_ds = New DataSet
		PCAD_BOM_ds.Tables.Add(DataTable_PCAD_Compare)
		DGV_PCAD_BOM.DataSource = Nothing
		DGV_PCAD_BOM.DataSource = DataTable_PCAD_Compare
		DGV_PCAD_BOM.AutoResizeColumns(DataGridViewAutoSizeColumnMode.AllCells)

		QB_BOM_ds = New DataSet
		QB_BOM_ds.Tables.Add(DataTable_QB_Compare)
		DGV_QB_BOM.DataSource = Nothing
		DGV_QB_BOM.DataSource = DataTable_QB_Compare
		DGV_QB_BOM.Visible = True
		DGV_QB_BOM.AutoResizeColumns(DataGridViewAutoSizeColumnMode.AllCells)

		ResizeTables(TP_BOM_compare)
	End Sub

	Private Sub B_Compare_ALPHA_Click() Handles B_Compare_ALPHA.Click
		Cursor = Cursors.WaitCursor
		If LOGDATA = True Then
			Try
				If ChangeCheck(True) = True Then
					CompareALPHABOM()
				End If
			Catch ex As Exception
				UnhandledExceptionMessage(ex)
			End Try
		Else
			If ChangeCheck(True) = True Then
				CompareALPHABOM()
			End If
		End If
		Cursor = Cursors.Default
	End Sub

	Private Sub CompareALPHABOM()
		Dim message As String = ""
		If sqlapi.CheckDirtyBit(message) = True Then
			MsgBox(message)
			Return
		End If

		TB_SearchIndicator.BackColor = Color.Black

		'Set up the tables that we are going to be using.
		Dim DataTable_PCAD_Compare As New DataTable
		DataTable_PCAD_Compare.Columns.Add(DB_HEADER_REF_DES, GetType(String))
		DataTable_PCAD_Compare.Columns.Add(DB_HEADER_ITEM_NUMBER, GetType(String))

		Dim DataTable_ALPHA_Compare As New DataTable
		DataTable_ALPHA_Compare.Columns.Add(DB_HEADER_REF_DES, GetType(String))
		DataTable_ALPHA_Compare.Columns.Add(DB_HEADER_ITEM_NUMBER, GetType(String))

		Dim board As String = ""
		Dim table As String = ""
		Dim optionValue As String = ""

		'Figure out if we are comparing a database entry or a search.
		If fromPCADdatabase Then
			'Create and add our query into our data set for comparison.
			PCAD_BOM_myCmd.CommandText = "SELECT [" & DB_HEADER_REF_DES & "], [" & DB_HEADER_ITEM_NUMBER & "], [" & DB_HEADER_OPTION & "] FROM " & TABLE_PCADBOM & " WHERE [" & DB_HEADER_BOARD_NAME & "] = '" & CB_Boards.Text & "' AND [" & DB_HEADER_PROCESS & "] = '" & PROCESS_SMT & "' ORDER BY [" & DB_HEADER_REF_DES & "]"
			board = CB_Boards.Text
			table = TABLE_PCADBOM
		ElseIf fromSearch Then
			'Create and add our query into our data set for comparison.
			Dim fileInformation As New FileInfo(My.Settings.BOMFilePath)
			Dim fileNameParsed() As String = fileInformation.Name.Split(".")
			If fileNameParsed.Length < 4 Then
				MsgBox("The file you have selected does not parse into 5 parts (boardname).(Rev#).(#).(Option).(bom)" & vbNewLine & "Please check the name of your file for correct naming conventions.")
				Return
			End If
			Dim fileName As String = fileNameParsed(INDEX_BOARD) & "." & fileNameParsed(INDEX_REVISION1) & "." & fileNameParsed(INDEX_REVISION2) & "." & fileNameParsed(INDEX_OPTION) & "."
			optionValue = fileNameParsed(INDEX_OPTION)
			PCAD_BOM_myCmd.CommandText = "SELECT [" & DB_HEADER_REF_DES & "], [" & DB_HEADER_ITEM_NUMBER & "], [" & DB_HEADER_OPTION & "] FROM " & TABLE_TEMP_PCADBOM & " WHERE [" & DB_HEADER_BOARD_NAME & "] = '" & fileName & "' AND [" & DB_HEADER_PROCESS & "] = '" & PROCESS_SMT & "' AND [" & DB_HEADER_OPTION & "] != 'X' ORDER BY [" & DB_HEADER_REF_DES & "]"
			board = Path.GetFileName(fileName)
			table = TABLE_TEMP_PCADBOM
		End If

		PCAD_BOM_da = New SqlDataAdapter(PCAD_BOM_myCmd)
		PCAD_BOM_ds = New DataSet()
		PCAD_BOM_da.Fill(PCAD_BOM_ds, "PCAD DATA")

		ALPHA_BOM_myCmd.commandText = "SELECT [" & DB_HEADER_REF_DES & "], [" & DB_HEADER_ITEM_NUMBER & "] FROM " & TABLE_ALPHABOM & " WHERE [" & DB_HEADER_BOARD_NAME & "] = '" & board & "' ORDER BY [" & DB_HEADER_REF_DES & "]"
		ALPHA_BOM_da = New SqlDataAdapter(ALPHA_BOM_myCmd)
		ALPHA_BOM_ds = New DataSet()
		ALPHA_BOM_da.Fill(ALPHA_BOM_ds, "ALPHA BOM DATA")

		If ALPHA_BOM_ds.Tables("ALPHA BOM DATA").Rows.Count <> 0 Then
			'-----------------------------------'
			'Compare ALPHA with our PCAD record '
			'-----------------------------------'

			For Each dsRow As DataRow In PCAD_BOM_ds.Tables("PCAD DATA").Rows
				'Check to see if we are dealing with 'ZD'.
				If dsRow(DB_HEADER_REF_DES).Contains(REFERENCE_DESIGNATOR_OPTION) = True Or dsRow(DB_HEADER_REF_DES).Contains(REFERENCE_DESIGNATOR_FIDUCIAL) = True Then
					Continue For
				End If

				Dim optionColumn As String = ""

				If fromSearch = True Then
					optionColumn = dsRow(DB_HEADER_OPTION)
				End If

				Dim include As Boolean = True
				'Check to see if we have an option field to account for from our file search.
				If optionColumn.Length <> 0 And fromSearch = True Then
					For index = 0 To optionColumn.Length - 1
						If optionValue.Contains(optionColumn(index)) = False Then
							include = False
							Continue For
						End If
					Next
				End If

				If include = False Then
					Continue For
				End If

				Dim ALPHABOMfound() As DataRow = ALPHA_BOM_ds.Tables("ALPHA BOM DATA").Select("[" & DB_HEADER_ITEM_NUMBER & "] = '" & dsRow(DB_HEADER_ITEM_NUMBER) & "' AND [" & DB_HEADER_REF_DES & "] = '" & dsRow(DB_HEADER_REF_DES) & "'")

				If ALPHABOMfound.Length = 0 Then
					'The item is not in the QB BOM.
					DataTable_PCAD_Compare.Rows.Add(dsRow(DB_HEADER_REF_DES), dsRow(DB_HEADER_ITEM_NUMBER))
				End If
			Next

			'-----------------------------------'
			'Compare PCAD with our ALPHA record '
			'-----------------------------------'
			For Each dsRow As DataRow In ALPHA_BOM_ds.Tables("ALPHA BOM DATA").Rows
				Dim PCADBOMfound() As DataRow = PCAD_BOM_ds.Tables("PCAD DATA").Select("[" & DB_HEADER_ITEM_NUMBER & "] = '" & dsRow(DB_HEADER_ITEM_NUMBER) & "' AND [" & DB_HEADER_REF_DES & "] = '" & dsRow(DB_HEADER_REF_DES) & "'")

				If PCADBOMfound.Length = 0 Then
					'The item is not in the QB BOM.
					DataTable_ALPHA_Compare.Rows.Add(dsRow(DB_HEADER_REF_DES), dsRow(DB_HEADER_ITEM_NUMBER))
				End If
			Next
		Else
			DataTable_PCAD_Compare.Rows.Add("There is no ALPHA BOM for: " & board)
			DataTable_ALPHA_Compare.Rows.Add("There is no ALPHA BOM for: " & board)
		End If

		fromCompareBOM = True
		fromCompareItems = False

		If DataTable_PCAD_Compare.Rows.Count = 0 And DataTable_ALPHA_Compare.Rows.Count = 0 Then
			DataTable_PCAD_Compare.Rows.Add("There are no missing components in the ALPHA BOM.")
			DataTable_ALPHA_Compare.Rows.Add("There are no extra components in the ALPHA BOM.")
		End If

		PCAD_BOM_ds = New DataSet
		PCAD_BOM_ds.Tables.Add(DataTable_PCAD_Compare)
		DGV_PCAD_BOM.DataSource = Nothing
		DGV_PCAD_BOM.DataSource = DataTable_PCAD_Compare
		DGV_PCAD_BOM.AutoResizeColumns(DataGridViewAutoSizeColumnMode.AllCells)

		'Use QB_BOM_ds for the create Excel Option
		QB_BOM_ds = New DataSet
		QB_BOM_ds.Tables.Add(DataTable_ALPHA_Compare)
		DGV_QB_BOM.DataSource = Nothing
		DGV_QB_BOM.DataSource = DataTable_ALPHA_Compare
		DGV_QB_BOM.Visible = True
		DGV_QB_BOM.AutoResizeColumns(DataGridViewAutoSizeColumnMode.AllCells)

		ResizeTables(TP_BOM_compare)
	End Sub

	Sub GetBoardDropDownItems(ByRef box As ComboBox)
		Dim BoardNames As New DataTable()
		Dim myCmd As New SqlCommand("SELECT Distinct([" & DB_HEADER_BOARD_NAME & "]) FROM " & TABLE_PCADBOM & " ORDER BY [" & DB_HEADER_BOARD_NAME & "]", myConn)

		BoardNames.Load(myCmd.ExecuteReader)

		For Each dr As DataRow In BoardNames.Rows
			box.Items.Add(dr(DB_HEADER_BOARD_NAME))
		Next

		If box.Items.Count <> 0 Then
			box.SelectedIndex = 0
		End If
		box.DropDownHeight = 200
	End Sub

	Private Sub CB_Boards_SelectedValueChanged() Handles CB_Boards.SelectedValueChanged
		B_CompareQBItems.Enabled = False
		B_CompareQBBOM.Enabled = False
		B_Compare_ALPHA.Enabled = False
	End Sub

	Private Sub DGV_PCAD_BOM_RowHeaderMouseDoubleClick(ByVal sender As Object, ByVal e As DataGridViewCellMouseEventArgs) Handles DGV_PCAD_BOM.RowHeaderMouseDoubleClick
		If fromCompareItems = True Then
			CompareFields(e.RowIndex, DGV_PCAD_BOM, DGV_QB_BOM)
		End If
	End Sub

	Private Sub DGV_QB_BOM_RowHeaderMouseDoubleClick(ByVal sender As Object, ByVal e As DataGridViewCellMouseEventArgs) Handles DGV_QB_BOM.RowHeaderMouseDoubleClick
		If fromCompareItems = True Then
			CompareFields(e.RowIndex, DGV_QB_BOM, DGV_PCAD_BOM)
		End If
	End Sub

	Private Sub CompareFields(ByRef rowIndex As Integer, ByRef clickedGridview As DataGridView, ByRef checkGridview As DataGridView)
		Dim message As String = ""
		If sqlapi.CheckDirtyBit(message) = True Then
			MsgBox(message)
			Return
		End If

		Dim PCADrowIndex As Integer
		Dim QBrowIndex As Integer

		'First we want to check to see if we even have the item in both BOMs
		Dim found As Boolean = False
		Dim foundIndex As Integer
		For Each row As DataGridViewRow In checkGridview.Rows
			If row.Cells.Item(DB_HEADER_ITEM_NUMBER).Value = clickedGridview.Rows(rowIndex).Cells.Item(DB_HEADER_ITEM_NUMBER).Value Then
				found = True
				foundIndex = row.Index
				Exit For
			End If
		Next

		If Not found Then
			MsgBox("The item number: " & clickedGridview.Rows(rowIndex).Cells.Item(DB_HEADER_ITEM_NUMBER).Value & " is not found in the other BOM.")
			Return
		Else
			If clickedGridview.Name = DGV_PCAD_BOM.Name Then
				PCADrowIndex = rowIndex
				QBrowIndex = foundIndex
			Else
				PCADrowIndex = foundIndex
				QBrowIndex = rowIndex
			End If
		End If

		Try
			Dim errorPresent As Boolean = False
			Dim errorList As New List(Of String)

			Dim pcadPrefix As String = DGV_PCAD_BOM.Rows(PCADrowIndex).Cells(DB_HEADER_ITEM_PREFIX).Value.ToString
			Dim pcadStockNumber As String = DGV_PCAD_BOM.Rows(PCADrowIndex).Cells(DB_HEADER_ITEM_NUMBER).Value.ToString

			Dim qbPrefix As String = DGV_QB_BOM.Rows(QBrowIndex).Cells(DB_HEADER_ITEM_PREFIX).Value.ToString
			Dim qbStockNumber As String = DGV_QB_BOM.Rows(QBrowIndex).Cells(DB_HEADER_ITEM_NUMBER).Value.ToString

			If String.Compare(pcadPrefix, qbPrefix) <> 0 Then
				errorList.Add("ITEM PREFIX")
				Dim pcad As String = "pcad: " & pcadPrefix
				Dim qb As String = "qb:   " & qbPrefix
				errorList.Add(pcad)
				errorList.Add(qb)
				errorList.Add("")

				errorPresent = True
			End If
			If String.Compare(pcadStockNumber, qbStockNumber) <> 0 Then
				errorList.Add("ITEM NUMBER")
				Dim pcad As String = "pcad: " & pcadStockNumber
				Dim qb As String = "qb:   " & qbStockNumber
				errorList.Add(pcad)
				errorList.Add(qb)
				errorList.Add("")

				errorPresent = True
			End If

			If CkB_Include.Checked = True Then
				Dim pcadVendor As String = DGV_PCAD_BOM.Rows(PCADrowIndex).Cells(DB_HEADER_VENDOR).Value.ToString
				Dim pcadPartNumber As String = DGV_PCAD_BOM.Rows(PCADrowIndex).Cells(DB_HEADER_MPN).Value.ToString

				Dim qbVendor As String = DGV_QB_BOM.Rows(QBrowIndex).Cells(DB_HEADER_VENDOR).Value.ToString
				Dim qbPartNumber As String = DGV_QB_BOM.Rows(QBrowIndex).Cells(DB_HEADER_MPN).Value.ToString

				If String.Compare(pcadVendor, qbVendor) <> 0 Then
					errorList.Add("VENDOR")
					Dim pcad As String = "pcad: " & pcadVendor
					Dim qb As String = "qb:   " & qbVendor
					errorList.Add(pcad)
					errorList.Add(qb)
					errorList.Add("")

					errorPresent = True
				End If
				If String.Compare(pcadPartNumber, qbPartNumber) <> 0 Then
					errorList.Add("MPN")
					Dim pcad As String = "pcad: " & pcadPartNumber
					Dim qb As String = "qb:   " & qbPartNumber
					errorList.Add(pcad)
					errorList.Add(qb)
					errorList.Add("")

					errorPresent = True
				End If
			End If

			If errorPresent = False Then
				MsgBox("There are no differences with the selected row")
			Else
				Dim messages As New MessageboxDifference(errorList)
				messages.Show()
			End If
		Catch ex As Exception
			MsgBox("There is no row to compare the selected row to.")
		End Try
	End Sub

	Private Sub DGV_PCAD_BOM_RowPostPaint(ByVal sender As Object, ByVal e As DataGridViewRowPostPaintEventArgs) Handles DGV_PCAD_BOM.RowPostPaint
		'Go through each row of the DGV and add the row number to the row header.
		Using b As SolidBrush = New SolidBrush(DGV_PCAD_BOM.RowHeadersDefaultCellStyle.ForeColor)
			e.Graphics.DrawString(e.RowIndex + 1, DGV_PCAD_BOM.DefaultCellStyle.Font, b, e.RowBounds.Location.X + 10, e.RowBounds.Location.Y + 4)
		End Using
	End Sub

	Private Sub DGV_QB_BOM_RowPostPaint(ByVal sender As Object, ByVal e As DataGridViewRowPostPaintEventArgs) Handles DGV_QB_BOM.RowPostPaint
		'Go through each row of the DGV and add the row number to the row header.
		Using b As SolidBrush = New SolidBrush(DGV_QB_BOM.RowHeadersDefaultCellStyle.ForeColor)
			e.Graphics.DrawString(e.RowIndex + 1 + QBitems_scrollValue, DGV_QB_BOM.DefaultCellStyle.Font, b, e.RowBounds.Location.X + 10, e.RowBounds.Location.Y + 4)
		End Using
	End Sub

	Private Sub DGV_PCAD_BOM_Scroll() Handles DGV_PCAD_BOM.Scroll
		Try
			DGV_QB_BOM.FirstDisplayedScrollingRowIndex = DGV_PCAD_BOM.FirstDisplayedScrollingRowIndex
		Catch ex As Exception

		End Try
	End Sub

	Private Sub DGV_QB_BOM_Scroll() Handles DGV_QB_BOM.Scroll
		Try
			DGV_PCAD_BOM.FirstDisplayedScrollingRowIndex = DGV_QB_BOM.FirstDisplayedScrollingRowIndex
		Catch ex As Exception

		End Try
	End Sub

	Private Sub B_Processes_Click() Handles B_Processes.Click
		Dim newForm As New PCADProcess()
		Dim frmCollection = Application.OpenForms
		For i = 0 To frmCollection.Count - 1
			If frmCollection.Item(i).Name = newForm.Name Then
				frmCollection.Item(i).Activate()
				frmCollection.Item(i).BringToFront()
				Exit Sub
			End If
		Next i
		newForm.Show()
	End Sub
#End Region

#Region "Tab 3: Release"
	Private Sub B_SearchSource_Click() Handles B_SearchSource.Click
		B_DeleteAlpha.Enabled = False
		B_Check.Enabled = True

		Dim selectFolder As New OpenFileDialog
		selectFolder.ValidateNames = False
		selectFolder.CheckFileExists = False
		selectFolder.CheckPathExists = True
		selectFolder.InitialDirectory = My.Settings.SourceFolderPath

		If selectFolder.ShowDialog() = DialogResult.OK Then
			My.Settings.SourceFolderPath = selectFolder.FileName.Substring(0, selectFolder.FileName.LastIndexOf("\"))
			My.Settings.Save()

			TB_SourceFolderPath.Text = My.Settings.SourceFolderPath
			LoadSourceGrid(My.Settings.SourceFolderPath)
			SearchForPartner()
		End If
	End Sub

	Private Sub B_SearchDestination_Click() Handles B_SearchDestination.Click
		Dim selectFolder As New OpenFileDialog
		selectFolder.ValidateNames = False
		selectFolder.CheckFileExists = False
		selectFolder.CheckPathExists = True
		selectFolder.InitialDirectory = My.Settings.DestinationFolderPath

		If selectFolder.ShowDialog() = DialogResult.OK Then
			My.Settings.DestinationFolderPath = selectFolder.FileName.Substring(0, selectFolder.FileName.LastIndexOf("\"))
			My.Settings.Save()

			TB_DestinationFolderPath.Text = My.Settings.DestinationFolderPath
			LoadDestinationGrid()
		End If
	End Sub

	Private Sub B_Add_Click() Handles B_Add.Click
		'First check to see if the files are read only.
		If CkB_ReadOnly.Checked = True Then
			MsgBox("Read Only files is checked.")
			Return
		End If

		'Second, make sure that we have at least one item selected from the list to copy over.
		If DGV_Source.SelectedRows.Count = 0 Then
			MsgBox("Please select at least one file from the source list.")
		Else
			For Each row In DGV_Source.SelectedRows
				Dim found As Boolean = False
				For i As Integer = 0 To DGV_Destination.RowCount - 1
					If DGV_Destination.Rows(i).Cells(0).Value.ToString = row.Cells(0).Value.ToString Then
						Dim result As Integer = MessageBox.Show("Do you want to replace duplicate files?", "Confirm Copy", MessageBoxButtons.YesNo)
						If result = DialogResult.No Then
							Exit Sub
						ElseIf result = DialogResult.Yes Then
							found = True
							Exit For
						End If
					End If
				Next
				If found = True Then
					Exit For
				End If
			Next

			'Look through for '+' symbol to tell us that it is a folder that we are going to copy over.
			For Each row In DGV_Source.SelectedRows
				If row.Cells(0).Value.ToString.Contains("+") = True Then
					CopyDirectory(My.Settings.SourceFolderPath, My.Settings.DestinationFolderPath, row.Cells(0).Value.ToString.Substring(2))
				Else
					If B_DeleteAlpha.Enabled = False Then
						CopyFile(My.Settings.SourceFolderPath, My.Settings.DestinationFolderPath, row.Cells(0).Value.ToString)
					Else
						CopyFile(ALPHA_BACKUP, My.Settings.DestinationFolderPath, row.Cells(0).Value.ToString)
					End If

				End If
			Next

			LoadDestinationGrid()
		End If
	End Sub

	Private Sub B_DeleteDestination_Click() Handles B_DeleteDestination.Click
		'First check to see if the files are read only.
		If CkB_ReadOnly.Checked = True Then
			MsgBox("Read Only files is checked.")
			Return
		End If

		'Second, make sure that we have at least one item selected from the list to copy over.
		If DGV_Destination.SelectedRows.Count = 0 Then
			MsgBox("Please select at least one file from the destination list.")
		Else
			Dim result As Integer = MessageBox.Show("Are you sure you want to delete the selected files in the Destination Folder?", "Confirm Delete", MessageBoxButtons.YesNo)
			If result = DialogResult.No Then
				Exit Sub
			End If

			'Look through for '+' symbol to tell us that it is a folder that we are going to copy over.
			For Each row In DGV_Destination.SelectedRows
				If row.cells(0).Value.ToString.Contains("+") = True Then
					Directory.Delete(My.Settings.DestinationFolderPath & "\" & row.Cells(0).Value.ToString.Substring(2), True)
				Else
					File.Delete(My.Settings.DestinationFolderPath & "\" & row.Cells(0).Value.ToString)
				End If
			Next

			LoadDestinationGrid()
		End If
	End Sub

	Private Function CheckforRelease(ByRef folderPath As String, ByRef datatable As DataTable, ByRef optionList As List(Of String), ByRef fromSource As Boolean) As Boolean
		Try
			Dim missingFilesList As New List(Of String)
			optionList = New List(Of String)
			Dim allbomFiles() As String = Directory.GetFiles(folderPath, PCAD_EXE)
			Dim name As String = ""
			Dim hasBOM As Boolean = False
			Dim needsPNP As Boolean = False
			Dim hasPNP As Boolean = False

			'Check for .BOM File
			If allbomFiles.Length = 0 Then
				datatable.Rows.Add()
				datatable.Rows(datatable.Rows.Count - 1)("Name") = "Missing .BOM.CSV"
				If fromSource Then
					B_BuildOptions.Enabled = False
				End If
				Return False
			Else
				If fromSource Then
					B_BuildOptions.Enabled = True
				End If
				Dim fileInformation As New FileInfo(allbomFiles(0))
				Dim nameParsed() As String = fileInformation.Name.Split(".")
				If nameParsed.Length < 4 Then
					MsgBox("The file you have selected does not parse into 5 parts (boardname).(Rev#).(#).(Option).(bom)" & vbNewLine & "Please check the name of your file for correct naming conventions.")
					Return False
				End If
				name = nameParsed(INDEX_BOARD) & "." & nameParsed(INDEX_REVISION1) & "." & nameParsed(INDEX_REVISION2) & "."
			End If

			Dim bomFiles() As String = Directory.GetFiles(folderPath, name & PCAD_EXE)
			Dim pcbFiles() As String = Directory.GetFiles(folderPath, name & PCB_EXE)
			Dim pnpFiles() As String = Directory.GetFiles(folderPath, name & PNP_CSV_EXE)
			Dim schFiles() As String = Directory.GetFiles(folderPath, name & SCH_EXE)
			Dim schpdfFiles() As String = Directory.GetFiles(folderPath, name & SCH_PDF_EXE)
			Dim revtxtFiles() As String = Directory.GetFiles(folderPath, name & REV_TXT_EXE)

			'Get the list of options that are valid to check for
			If bomFiles.Length = 0 Then
				missingFilesList.Add("Missing BOM.CSV")
			Else
				hasBOM = True
				Using myParser As New TextFieldParser(bomFiles(0))
					'Indexs
					Dim INDEX_refdes As Integer = -1
					Dim INDEX_value As Integer = -1
					Dim INDEX_process As Integer = -1
					'Optional
					Dim INDEX_option As Integer = -1

					myParser.TextFieldType = FieldType.Delimited
					myParser.SetDelimiters(",")
					Dim currentRow As String()

					'First row is the header row.
					currentRow = myParser.ReadFields()
					Dim index As Integer = 0

					'Parse the header row to grab Indexs. They can be generated in any order.
					For Each header In currentRow
						Select Case header.ToLower
							Case "refdes"
								INDEX_refdes = index
							Case "value"
								INDEX_value = index
							Case "option"
								INDEX_option = index
							Case "process"
								INDEX_process = index
						End Select
						index += 1
					Next

					While Not myParser.EndOfData
						currentRow = myParser.ReadFields()

						If currentRow(0).Length = 0 Then
							Continue While
						End If

						'Parse reference designator.
						If currentRow(INDEX_refdes).Contains(REFERENCE_DESIGNATOR_OPTION) = True Then
							If currentRow(INDEX_option).Contains("X") = True Then
								Continue While
							End If

							optionList.Add(currentRow(INDEX_option))
						End If

						If String.Compare(currentRow(INDEX_process), PROCESS_SMT, True) = 0 Then
							needsPNP = True
						End If
					End While
				End Using
			End If

			'Check to see that each option has its own BOM
			If optionList.Count <> 0 Then
				If fromSource = True Then
					B_BuildOptions.Enabled = True
				End If

				For Each item In optionList
					Dim found As Boolean = False
					For Each File In allbomFiles
						If File.Contains(name & item & ".") = True Then
							found = True
						End If
					Next
					If found = False Then
						missingFilesList.Add("Missing " & item & " BOM.CSV")
					End If
				Next
			Else
				If fromSource = True Then
					B_BuildOptions.Enabled = False
				End If
			End If

			'Check for .PCB File
			If pcbFiles.Length = 0 Then
				missingFilesList.Add("Missing .PCB")
			ElseIf pcbFiles.Length > 1 Then
				missingFilesList.Add("Extra .PCB")
			End If

			'Check for .PNP.CSV File
			If needsPNP = True Then
				If pnpFiles.Length = 0 Then
					missingFilesList.Add("Missing .PNP.CSV")
				ElseIf pnpFiles.Length > 1 Then
					missingFilesList.Add("Extra .PNP.CSV")
				Else
					hasPNP = True
				End If
			Else
				hasPNP = True
			End If

			'Check for .SCH File
			If schFiles.Length = 0 Then
				missingFilesList.Add("Missing .SCH")
			ElseIf schFiles.Length > 1 Then
				missingFilesList.Add("Extra .SCH")
			End If

			'Check for .SCH.PDF File
			If schpdfFiles.Length = 0 Then
				missingFilesList.Add("Missing .SCH.PDF")
			ElseIf schpdfFiles.Length > 1 Then
				missingFilesList.Add("Extra .SCH.PDF")
			End If

			'Check for .ZIP File
			Try
				Dim zipFiles() As String = Directory.GetFiles(folderPath & "\Released", name & ZIP_EXE)
				If zipFiles.Length = 0 Then
					missingFilesList.Add("Missing .ZIP")
				ElseIf zipFiles.Length > 1 Then
					missingFilesList.Add("Extra .ZIP")
				End If
			Catch ex As Exception
				missingFilesList.Add("Missing Released Folder")
				missingFilesList.Add("Missing .ZIP")
			End Try

			'Check for .REV.TXT File
			If revtxtFiles.Length = 0 Then
				missingFilesList.Add("Missing .REV.TXT")
			ElseIf revtxtFiles.Length > 1 Then
				missingFilesList.Add("Extra .REV.TXT")
			End If

			If missingFilesList.Count = 0 Then
				If fromSource = True And hasBOM = True And hasPNP = True And needsPNP = True Then
					If CompareBOMandPNP(datatable, bomFiles(0), pnpFiles(0)) = False Then
						Return False
					Else
						Return True
					End If
				Else
					Return True
				End If
			Else
				For Each item In missingFilesList
					datatable.Rows.Add()
					datatable.Rows(datatable.Rows.Count - 1)("Name") = item
				Next
				If fromSource = True And hasBOM = True And hasPNP = True And needsPNP = True Then
					CompareBOMandPNP(datatable, bomFiles(0), pnpFiles(0))
				End If
				Return False
			End If
		Catch ex As Exception
			MsgBox(ex.Message)
			Return False
		End Try
	End Function

	Private Function CompareBOMandPNP(ByRef DataTable As DataTable, ByRef bomPath As String, ByRef pnpPath As String) As Boolean
		Try
			Dim issueList As New List(Of String)

			If ParseBOMFile(bomPath, True) = False Then
				Return False
			End If
			If ParsePNPFile(pnpPath, issueList) = False Then
				For Each item In issueList
					DataTable.Rows.Add()
					DataTable.Rows(DataTable.Rows.Count - 1)("Name") = item
				Next
				Return False
			End If

			Dim PCAD_Temp = New DataTable
			myCmd.CommandText = "SELECT * FROM " & TABLE_TEMP_PCADBOM & " WHERE [" & DB_HEADER_PROCESS & "] = '" & PROCESS_SMT & "' ORDER BY [" & DB_HEADER_REF_DES & "]"
			PCAD_Temp.Load(myCmd.ExecuteReader())

			Dim PNP_Temp = New DataTable
			myCmd.CommandText = "SELECT * FROM " & TABLE_TEMP_PNP & " ORDER BY [" & DB_HEADER_REF_DES & "]"
			PNP_Temp.Load(myCmd.ExecuteReader())

			'-------------------------------------------------------'
			'                                                       '
			'   Check for components that are missing in the PNP.   '
			'                                                       '
			'-------------------------------------------------------'
			For Each dsRow As DataRow In PCAD_Temp.Rows
				If dsRow(DB_HEADER_REF_DES).ToString.Contains(REFERENCE_DESIGNATOR_SWAP) = False Then
					'We are not dealing with a swap so we can use the original Reference designator.
					If PNP_Temp.Select("[" & DB_HEADER_ITEM_NUMBER & "] = '" & dsRow(DB_HEADER_ITEM_NUMBER) & "' AND [" & DB_HEADER_REF_DES & "] = '" & dsRow(DB_HEADER_REF_DES) & "' AND [" & DB_HEADER_ITEM_PREFIX & "] = '" & dsRow(DB_HEADER_ITEM_PREFIX) & "'").Length = 0 Then
						'The Item Number is not in the PNP database.
						issueList.Add("PNP Missing " & dsRow(DB_HEADER_REF_DES))
					End If
				Else
					If dsRow(DB_HEADER_SWAP).ToString.Length <> 0 Then
						If PNP_Temp.Select("[" & DB_HEADER_REF_DES & "] = '" & dsRow(DB_HEADER_SWAP) & "'").Length = 0 Then
							'The Item Number is not in the PNP database.
							issueList.Add("BOM Swap " & dsRow(DB_HEADER_REF_DES))
						End If
					End If
				End If
			Next

			For Each dsRow As DataRow In PNP_Temp.Rows
				If String.Compare(dsRow(DB_HEADER_PROCESS), PROCESS_SMT) <> 0 Then
					If String.Compare(dsRow(DB_HEADER_PROCESS), PROCESS_NOTUSED) <> 0 Then
						issueList.Add("PNP " & dsRow(DB_HEADER_REF_DES) & " Process")
					End If
					Continue For
				End If

				If PCAD_Temp.Select("[" & DB_HEADER_ITEM_NUMBER & "] = '" & dsRow(DB_HEADER_ITEM_NUMBER) & "' AND [" & DB_HEADER_REF_DES & "] = '" & dsRow(DB_HEADER_REF_DES) & "' AND [" & DB_HEADER_ITEM_PREFIX & "] = '" & dsRow(DB_HEADER_ITEM_PREFIX) & "'").Length = 0 Then
					'The Item Number is not in the PNP database.
					issueList.Add("PNP Extra " & dsRow(DB_HEADER_REF_DES))
				End If
			Next

			If issueList.Count = 0 Then
				Return True
			Else
				For Each item In issueList
					DataTable.Rows.Add()
					DataTable.Rows(DataTable.Rows.Count - 1)("Name") = item
				Next
				Return False
			End If
		Catch ex As Exception
			MsgBox(ex.Message)
			Return False
		End Try
	End Function

	Private Function ParsePNPFile(ByRef filePath As String, ByRef issueList As List(Of String)) As Boolean
		Dim myCmd As New SqlCommand("DELETE FROM " & TABLE_TEMP_PNP, myConn)
		myCmd.ExecuteNonQuery()

		'Indexs
		Dim INDEX_refdes As Integer = -1
		Dim INDEX_stockNumber As Integer = -1
		Dim INDEX_PosX As Integer = -1
		Dim INDEX_PosY As Integer = -1
		Dim INDEX_Rotation As Integer = -1
		Dim INDEX_Process As Integer = -1
		Dim INDEX_Value As Integer = -1
		'Optional
		Dim INDEX_option As Integer = -1
		Dim INDEX_swap As Integer = -1

		Dim br_check As Integer = -1
		Dim fn_check As Integer = -1
		Dim fu1_check As Integer = -1
		Dim fu2_check As Integer = -1
		Dim fu3_check As Integer = -1

		'Start our transaction. Must assign both transaction object and connection to the command object for a pending local transaction.
		Dim transaction As SqlTransaction = Nothing
		transaction = myConn.BeginTransaction("Temp Transaction")
		myCmd.Connection = myConn
		myCmd.Transaction = transaction

		Try
			Using myParser As New TextFieldParser(filePath)
				myParser.TextFieldType = FieldType.Delimited
				myParser.SetDelimiters(",")
				Dim currentRow As String()

				'First row is the header row.
				currentRow = myParser.ReadFields()
				currentRow = myParser.ReadFields()
				currentRow = myParser.ReadFields()
				Dim index As Integer = 0
				Dim missingFields As String = ""
				Dim fieldErrors As Boolean = False

				'Parse the header row to grab Indexs. They can be generated in any order.
				For Each header In currentRow
					Select Case header.ToLower
						Case "refdes"
							INDEX_refdes = index
						Case "stock number"
							INDEX_stockNumber = index
						Case "locationx"
							INDEX_PosX = index
						Case "locationy"
							INDEX_PosY = index
						Case "rotation"
							INDEX_Rotation = index
						Case "option"
							INDEX_option = index
						Case "process"
							INDEX_Process = index
						Case "swap"
							INDEX_swap = index
						Case "value"
							INDEX_Value = index
					End Select
					index += 1
				Next

				If INDEX_refdes = -1 Then
					fieldErrors = True
					missingFields = missingFields & "refdes |"
				End If
				If INDEX_stockNumber = -1 Then
					fieldErrors = True
					missingFields = missingFields & " stockNumber |"
				End If
				If INDEX_PosX = -1 Then
					fieldErrors = True
					missingFields = missingFields & " locationX |"
				End If
				If INDEX_PosY = -1 Then
					fieldErrors = True
					missingFields = missingFields & " locationY |"
				End If
				If INDEX_Rotation = -1 Then
					fieldErrors = True
					missingFields = missingFields & " Rotation |"
				End If
				If INDEX_Process = -1 Then
					fieldErrors = True
					missingFields = missingFields & " Process |"
				End If
				If INDEX_Value = -1 Then
					fieldErrors = True
					missingFields = missingFields & " Value"
				End If

				If fieldErrors = True Then
					issueList.Add("PNP File is missing the following Fields: " & missingFields)
					myParser.Close()
					transaction.Rollback()
					Return False
				End If

				While Not myParser.EndOfData
					Dim referenceDesignator As String = ""
					Dim prefix As String = ""
					Dim stockNumber As String = ""
					Dim posX As String = ""
					Dim posY As String = ""
					Dim rotation As String = ""
					Dim process As String = ""
					'Optional
					Dim optionValue As String = ""
					currentRow = myParser.ReadFields()

					If currentRow(0).Length = 0 Then
						Continue While
					End If

					'Parse the Reference Designator.
					referenceDesignator = currentRow(INDEX_refdes)

					Select Case referenceDesignator
						Case "FU1"
							fu1_check = 0
							If currentRow(INDEX_Value).Contains("|") = True Then
								Dim parsed() As String = currentRow(INDEX_Value).Split("|")
								For Each item In parsed
									Select Case item.Substring(0, 2).ToLower
										Case "br"
											br_check = 0
										Case "fn"
											fn_check = 0
									End Select
								Next
								Continue While
							Else
								issueList.Add("FU1 '|' format")
								myParser.Close()
								transaction.Rollback()
								Return False
							End If
						Case "FU2"
							fu2_check = 0
							Continue While
						Case "FU3"
							fu3_check = 0
							Continue While
					End Select


					'Check to see if we have the option field or not.
					If INDEX_option <> -1 Then
						optionValue = currentRow(INDEX_option)
					End If

					'Parse Stock Number.
					stockNumber = currentRow(INDEX_stockNumber).Substring(currentRow(INDEX_stockNumber).IndexOf(":") + 1)

					'Parse the Prefix.
					If currentRow(INDEX_stockNumber).Contains(":") = True Then
						prefix = currentRow(INDEX_stockNumber).Substring(0, currentRow(INDEX_stockNumber).IndexOf(":"))
					End If

					'Parse X position.
					posX = currentRow(INDEX_PosX).Substring(0, currentRow(INDEX_PosX).IndexOf("."))

					'Parse Y position.
					posY = currentRow(INDEX_PosY).Substring(0, currentRow(INDEX_PosY).IndexOf("."))

					'Parse Rotation.
					rotation = currentRow(INDEX_Rotation).Substring(0, currentRow(INDEX_Rotation).IndexOf("."))

					'Parse the Process.
					process = currentRow(INDEX_Process)

					myCmd.CommandText = "INSERT INTO " & TABLE_TEMP_PNP & " ([" & DB_HEADER_REF_DES & "], [" & DB_HEADER_ITEM_PREFIX & "], [" & DB_HEADER_ITEM_NUMBER & "], [" & DB_HEADER_POS_X & "], [" & DB_HEADER_POS_Y & "], [" & DB_HEADER_ROTATION & "], [" & DB_HEADER_PROCESS & "]) " &
										"VALUES('" & referenceDesignator & "', '" & prefix & "', '" & stockNumber & "', '" & posX & "', '" & posY & "', '" & rotation & "', '" & process & "')"
					myCmd.ExecuteNonQuery()
				End While
			End Using

			'Check to see if we have found each of our Fiducial information parts
			If br_check = -1 Or fn_check = -1 Or fu1_check = -1 Or fu2_check = -1 Or fu3_check = -1 Then
				Dim infostring As String = "Fiducial Information missing:"
				If br_check = -1 Then
					infostring = infostring & " BR"
				End If
				If fn_check = -1 Then
					infostring = infostring & " FN"
				End If
				If fu1_check = -1 Then
					infostring = infostring & " FU1"
				End If
				If fu2_check = -1 Then
					infostring = infostring & " FU2"
				End If
				If fu3_check = -1 Then
					infostring = infostring & " FU3"
				End If

				issueList.Add(infostring)
				transaction.Rollback()
				Return False
			End If
			transaction.Commit()
		Catch ex As Exception
			If Not transaction Is Nothing Then
				sqlapi.RollBack(transaction, errorMessage:=New List(Of String))
				MsgBox("COULD NOT FIND PNP FILE AT: " & filePath)
				Return False
			End If
		End Try
		Return True
	End Function

	Private Sub DGV_Source_CellValueChanged(ByVal sender As Object, ByVal e As DataGridViewCellEventArgs) Handles DGV_Source.CellValueChanged
		'Make sure that the cell that was changed is the New Name column. 
		If DGV_Source.Columns(e.ColumnIndex).Name = "New Name" Then
			Try
				'Figure out if we are renaming a folder or a file.
				If DGV_Source(0, e.RowIndex).Value.ToString.Contains("+") = True Then
					My.Computer.FileSystem.RenameDirectory(My.Settings.SourceFolderPath & "\" & DGV_Source(0, e.RowIndex).Value.ToString.Substring(2), DGV_Source(1, e.RowIndex).Value.ToString)
				Else
					My.Computer.FileSystem.RenameFile(My.Settings.SourceFolderPath & "\" & DGV_Source(0, e.RowIndex).Value.ToString, DGV_Source(1, e.RowIndex).Value.ToString)
				End If
			Catch ex As Exception
				MsgBox(ex.Message)
			End Try
		End If
		'Reload the directory with the new file/folder names.
		LoadSourceGrid(My.Settings.SourceFolderPath)
	End Sub

	Private Sub DGV_Source_CellDoubleClick(ByVal sender As Object, ByVal e As DataGridViewCellEventArgs) Handles DGV_Source.CellDoubleClick
		'Make sure we clicked a row that exists.
		If e.RowIndex <> -1 Then
			'Make sure that we clicked a a folder with the (+) symbol because those are folders.
			If DGV_Source(0, e.RowIndex).Value.ToString.Contains("+") = True Then
				TB_SourceFolderPath.Text = My.Settings.SourceFolderPath & "\" & DGV_Source(0, e.RowIndex).Value.ToString.Substring(2)

				TB_SourceIndicatorLight.BackColor = Color.Black
				B_CopyReleaseFiles.Enabled = False
				B_Back.Enabled = True
				B_Add.Enabled = False
				B_Check.Enabled = False

				LoadSourceGrid(My.Settings.SourceFolderPath & "\" & DGV_Source(0, e.RowIndex).Value.ToString.Substring(2))
			End If
		End If
	End Sub

	Private Sub B_Back_Click() Handles B_Back.Click
		TB_SourceFolderPath.Text = My.Settings.SourceFolderPath

		B_DeleteAlpha.Enabled = False
		B_Back.Enabled = False
		B_Add.Enabled = True
		B_Check.Enabled = True

		'Load up the previous folder location.
		LoadSourceGrid(My.Settings.SourceFolderPath)
	End Sub

	Private Sub CkB_ReadOnly_CheckedChanged() Handles CkB_ReadOnly.CheckedChanged
		'Check to see if we are checking or un-checking the box. 
		'This will determine if we are making all of the files/folders in the released directory ReadOnly or not.
		If CkB_ReadOnly.Checked = True Then
			For Each dir As String In Directory.GetDirectories(My.Settings.DestinationFolderPath)
				Dim fileInformation As New DirectoryInfo(dir)
				For Each item As FileInfo In fileInformation.GetFiles
					item.IsReadOnly = True
				Next
			Next
			For Each dir As String In Directory.GetFiles(My.Settings.DestinationFolderPath)
				Dim fileInformation As New FileInfo(dir)
				fileInformation.IsReadOnly = True
			Next
		Else
			For Each dir As String In Directory.GetDirectories(My.Settings.DestinationFolderPath)
				Dim fileInformation As New DirectoryInfo(dir)
				For Each item As FileInfo In fileInformation.GetFiles
					item.IsReadOnly = False
				Next
			Next
			For Each dir As String In Directory.GetFiles(My.Settings.DestinationFolderPath)
				Dim fileInformation As New FileInfo(dir)
				fileInformation.IsReadOnly = False
			Next
		End If
	End Sub

	Private Sub B_CopyReleaseFiles_Click() Handles B_CopyReleaseFiles.Click
		'First, check to see if the files are read only.
		If CkB_ReadOnly.Checked = True Then
			MsgBox("Read Only files is checked.")
			Return
		End If
		For Each row In DGV_Source.Rows
			Dim found As Boolean = False
			For i As Integer = 0 To DGV_Destination.RowCount - 1

				'Check to see if we have any fiels that are found in both locatins.
				If DGV_Destination.Rows(i).Cells(0).Value.ToString = row.Cells(0).Value.ToString Then
					Dim result As Integer = MessageBox.Show("Do you want to replace duplicate files?", "Confirm Copy", MessageBoxButtons.YesNo)
					If result = DialogResult.No Then
						Exit Sub
					ElseIf result = DialogResult.Yes Then
						found = True
						Exit For
					End If
				End If
			Next
			If found = True Then
				Exit For
			End If
		Next


		'Get the name of the .BOM File to use to find all of the other file names.
		Dim allbomFiles() As String = Directory.GetFiles(My.Settings.SourceFolderPath, PCAD_EXE)

		Dim fileInformation As New FileInfo(allbomFiles(0))
		Dim nameParsed() As String = fileInformation.Name.Split(".")
		If nameParsed.Length < 4 Then
			MsgBox("The file you have selected does not parse into 5 parts (boardname).(Rev#).(#).(Option).(bom)" & vbNewLine & "Please check the name of your file for correct naming conventions.")
			Return
		End If
		Name = nameParsed(INDEX_BOARD) & "." & nameParsed(INDEX_REVISION1) & "." & nameParsed(INDEX_REVISION2) & "."

		'Check to make sure that the file name is the same as the directory name.
		Dim locationParsed() As String = My.Settings.SourceFolderPath.Split("\")

		Dim nameissue As String = ""
		If nameParsed(INDEX_BOARD) <> locationParsed(locationParsed.Length - 2) Then
			nameissue = "Folder Location " & locationParsed(locationParsed.Length - 2) & " and file name " & nameParsed(INDEX_BOARD) & " do not match" & vbNewLine
		End If

		If nameParsed(INDEX_REVISION1) & "." & nameParsed(INDEX_REVISION2) <> locationParsed(locationParsed.Length - 1) Then
			nameissue = nameissue & "Folder Location " & locationParsed(locationParsed.Length - 1) & " and file revision " & nameParsed(INDEX_REVISION1) & "." & nameParsed(INDEX_REVISION2) & " do not match"
		End If

		If nameissue.Length <> 0 Then
			Dim answer = MessageBox.Show(nameissue & vbNewLine & "Would you like to continue?", "Continue?", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
			If answer = DialogResult.No Then
				Return
			End If
		End If


		'Copy the .PCB Files.
		Dim pcbFiles() As String = Directory.GetFiles(My.Settings.SourceFolderPath, Name & PCB_EXE)
		fileInformation = New FileInfo(pcbFiles(0))
		CopyFile(My.Settings.SourceFolderPath, My.Settings.DestinationFolderPath, fileInformation.Name)


		'Copy the .PNP.CSV Files.
		Dim pnpFiles() As String = Nothing
		Try
			pnpFiles = Directory.GetFiles(My.Settings.SourceFolderPath, Name & PNP_CSV_EXE)
			fileInformation = New FileInfo(pnpFiles(0))
			CopyFile(My.Settings.SourceFolderPath, My.Settings.DestinationFolderPath, fileInformation.Name)
		Catch ex As Exception

		End Try

		'Copy the .SCH Files.
		Dim schFiles() As String = Directory.GetFiles(My.Settings.SourceFolderPath, Name & SCH_EXE)
		fileInformation = New FileInfo(schFiles(0))
		CopyFile(My.Settings.SourceFolderPath, My.Settings.DestinationFolderPath, fileInformation.Name)


		'Copy the .SCH.PDF Files.
		Dim schpdfFiles() As String = Directory.GetFiles(My.Settings.SourceFolderPath, Name & SCH_PDF_EXE)
		fileInformation = New FileInfo(schpdfFiles(0))
		CopyFile(My.Settings.SourceFolderPath, My.Settings.DestinationFolderPath, fileInformation.Name)


		'Copy the .BOM.CSV Files.
		Dim bomFiles() As String = Directory.GetFiles(My.Settings.SourceFolderPath, Name & PCAD_EXE)
		If optionListSource.Count <> 0 Then
			For Each fileName In bomFiles
				Dim found As Boolean = False
				For Each item In optionListSource
					If fileName.Contains("." & item & ".") = True Then
						found = True
						Exit For
					End If
				Next
				If found = True Then
					fileInformation = New FileInfo(fileName)
					CopyFile(My.Settings.SourceFolderPath, My.Settings.DestinationFolderPath, fileInformation.Name)

					Dim nameParsed2() As String = fileInformation.Name.Split(".")
					If nameParsed2.Length < 4 Then
						MsgBox("The file you have selected does not parse into 5 parts (boardname).(Rev#).(#).(Option).(bom)" & vbNewLine & "Please check the name of your file for correct naming conventions.")
						Return
					End If
					Dim boardName As String = nameParsed2(INDEX_BOARD) & "." & nameParsed2(INDEX_REVISION1) & "." & nameParsed2(INDEX_REVISION2) & "." & nameParsed2(INDEX_OPTION) & "."

					CreatePNPList(pnpFiles(0), boardName)


					'Check to see if the Directory has been created yet. And then create the fake file.
					'ALPHA
					If Directory.Exists(RELEASE & "\ALPHABOM\" & locationParsed(locationParsed.Length - 2) & "\" & locationParsed(locationParsed.Length - 1)) = False Then
						Directory.CreateDirectory(RELEASE & "\ALPHABOM\" & locationParsed(locationParsed.Length - 2) & "\" & locationParsed(locationParsed.Length - 1))
					End If

					If File.Exists(RELEASE & "\ALPHABOM\" & locationParsed(locationParsed.Length - 2) & "\" & locationParsed(locationParsed.Length - 1) & "\" & boardName & ".gen") = False Then
						File.Create(RELEASE & "\ALPHABOM\" & locationParsed(locationParsed.Length - 2) & "\" & locationParsed(locationParsed.Length - 1) & "\" & boardName & ".gen")
					End If
				End If
			Next
		Else
			fileInformation = New FileInfo(bomFiles(0))

			Dim boardName As String = fileInformation.Name
			CopyFile(My.Settings.SourceFolderPath, My.Settings.DestinationFolderPath, boardName)

			Dim nameParsed2() As String = boardName.Split(".")
			If nameParsed2.Length < 4 Then
				MsgBox("The file you have selected does not parse into 5 parts (boardname).(Rev#).(#).(Option).(bom)" & vbNewLine & "Please check the name of your file for correct naming conventions.")
				Return
			End If
			boardName = nameParsed2(INDEX_BOARD) & "." & nameParsed2(INDEX_REVISION1) & "." & nameParsed2(INDEX_REVISION2) & "." & nameParsed2(INDEX_OPTION) & "."

			CreatePNPList(pnpFiles(0), boardName)

		End If

		'Copy the Release Folder.
		CopyDirectory(My.Settings.SourceFolderPath, My.Settings.DestinationFolderPath, "Released")

		'Copy the .REV.TXT Files.
		Dim revtxtFiles() As String = Directory.GetFiles(My.Settings.SourceFolderPath, Name & REV_TXT_EXE)

		' the file will be copied if we pass the UpdateRevisionFile()
		' add the contents of the .REV.TXT file to the master excel
		UpdateRevisionFile(revtxtFiles(0))

		'Double check and make sure we are still Release ready.
		LoadDestinationGrid()
	End Sub

	Private Sub UpdateRevisionFile(ByRef file As String)
		' first we need to check to see if we ahve a valid excel file in our settings
		If IO.File.Exists(My.Settings.RevisionFile) = False Then
			Using obj As New OpenFileDialog
				obj.Filter = "Excel|*.xlsx"
				obj.CheckFileExists = False
				obj.CheckPathExists = False
				obj.InitialDirectory = My.Settings.RevisionFile
				obj.CustomPlaces.Add("\\Server1")
				obj.CustomPlaces.Add("C:")
				obj.Title = "Select Revision File"

				If obj.ShowDialog = Windows.Forms.DialogResult.OK Then
					My.Settings.RevisionFile = obj.FileName
					My.Settings.Save()
				Else
					MsgBox("Revision file was not updated because the user decided not to locate it.")
					Return
				End If
			End Using
		End If

		Dim report As New GenerateReport

		Dim fileInformation As New FileInfo(file)
		Dim nameParsed() As String = fileInformation.Name.Split(".")
		If report.ModifyExcel(file, nameParsed(INDEX_BOARD)) = False Then
			Return
		End If

		' copy the file over after we pass the excel test so we can force the release fail if we do not update
		CopyFile(My.Settings.SourceFolderPath, My.Settings.DestinationFolderPath, fileInformation.Name)

	End Sub

	Private Sub B_BuildOptions_Click() Handles B_BuildOptions.Click
		'First check to see that we have the correct BOM.CSV file to build our options from.
		Dim bomFiles() As String = Directory.GetFiles(My.Settings.SourceFolderPath, PCAD_EXE)
		If bomFiles.Count = 0 Then
			MsgBox("Could not build options. Missing .BOM.CSV file.")
			Exit Sub
		End If
		Dim fileInformation As New FileInfo(bomFiles(0))
		Dim fileNameParsed() As String = fileInformation.Name.Split(".")
		If fileNameParsed.Length < 4 Then
			MsgBox("The file you have selected does not parse into 5 parts (boardname).(Rev#).(#).(Option).(bom)" & vbNewLine & "Please check the name of your file for correct naming conventions.")
			Return
		End If

		Dim result As Integer = MessageBox.Show("Would you like to overwirte any duplicate files?", "Overwrite", MessageBoxButtons.YesNo)
		If result = DialogResult.No Then
			Exit Sub
		End If

		'Go through the list of options that was made when we parsed through the BOM and create a new file with each option in the file name.
		For Each item In optionListSource
			Dim fileName As String = fileNameParsed(INDEX_BOARD) & "." & fileNameParsed(INDEX_REVISION1) & "." & fileNameParsed(INDEX_REVISION2) & "." & item & ".bom.csv"
			Try
				File.Delete(My.Settings.SourceFolderPath & "\" & fileName)
			Catch ex As Exception

			End Try
			File.Copy(My.Settings.SourceFolderPath & "\" & fileInformation.Name, My.Settings.SourceFolderPath & "\" & fileName, True)
			File.SetLastWriteTime(My.Settings.SourceFolderPath & "\" & fileName, Date.Now)
		Next

		'Re-Update our gridview
		LoadSourceGrid(My.Settings.SourceFolderPath)
	End Sub

	Private Sub B_ALPHA_Click() Handles B_ALPHA.Click
		B_DeleteAlpha.Enabled = True
		B_Back.Enabled = True
		B_Check.Enabled = False

		TB_SourceFolderPath.Text = ALPHA_BACKUP

		'Load up the ALPHA directory location.
		LoadSourceAlphaGrid()
	End Sub

	Private Sub B_DeleteAlpha_Click() Handles B_DeleteAlpha.Click
		'Make sure that we have selected one row to delete.
		If DGV_Source.SelectedRows.Count = 0 Then
			MsgBox("Please select at least one file from the Source list.")
		Else
			Dim result As Integer = MessageBox.Show("Are you sure you want to delete the selected files in the Source Folder?", "Confirm Delete", MessageBoxButtons.YesNo)
			If result = DialogResult.No Then
				Exit Sub
			End If

			For Each row In DGV_Source.SelectedRows
				File.Delete(ALPHA_BACKUP & "\" & row.Cells(0).Value.ToString)
			Next
			LoadSourceAlphaGrid()
		End If
	End Sub

	Private Sub SearchForPartner()
		Dim locationParsed() As String = My.Settings.SourceFolderPath.Split("\")

		'Check to see if the directory exists.
		'If not, then ask of the user would like to create it.
		If Directory.Exists(RELEASE_PCAD & "\" & locationParsed(locationParsed.Length - 2) & "\" & locationParsed(locationParsed.Length - 1)) = False Then
			Dim result As Integer = MessageBox.Show(My.Settings.SourceFolderPath & " does not have a release location already." & vbNewLine &
													"Would you like to create this folder directory for it?" & vbNewLine &
													locationParsed(locationParsed.Length - 2) & "  " & locationParsed(locationParsed.Length - 1), "Overwrite", MessageBoxButtons.YesNo)
			If result = DialogResult.No Then
				Exit Sub
			End If
			Directory.CreateDirectory(RELEASE_PCAD & "\" & locationParsed(locationParsed.Length - 2) & "\" & locationParsed(locationParsed.Length - 1))
		End If
		My.Settings.DestinationFolderPath = RELEASE_PCAD & "\" & locationParsed(locationParsed.Length - 2) & "\" & locationParsed(locationParsed.Length - 1)
		My.Settings.Save()

		TB_DestinationFolderPath.Text = My.Settings.DestinationFolderPath

		'Re-Update our gridview
		LoadDestinationGrid()
	End Sub

	Private Sub LoadSourceGrid(ByRef path As String)
		DataTable_source = New DataTable
		DataTable_source.Columns.Add("Name")
		DataTable_source.Columns.Add("New Name")
		DataTable_source.Columns.Add("Date")

		'When passing in a path, if it is a shared network drive you need to pass more of the path
		' \\Server1 should be \\Server1\Shares

		'Check to see if we are dealing with our ALPHA location.
		'If not, then we want to add all of the folders to the top of our DGV.
		If My.Settings.SourceFolderPath.ToLower.Contains("alpha") = False Then
			For Each dir As String In Directory.GetDirectories(path)
				Dim fileInformation As New FileInfo(dir)
				DataTable_source.Rows.Add("+ " & fileInformation.Name, fileInformation.Name, fileInformation.LastWriteTime)
			Next
		End If
		For Each dir As String In Directory.GetFiles(path)
			Dim fileInformation As New FileInfo(dir)
			DataTable_source.Rows.Add(fileInformation.Name, fileInformation.Name, fileInformation.LastWriteTime)
		Next

		'Check to see if we are no longer on the correct level for our 'Release Check'.
		'If not, then we need to make our indicator light not show that it is ready/not ready.
		If B_Back.Enabled = False Then
			If My.Settings.SourceFolderPath.ToLower.Contains("alpha") = False Then
				If CheckforRelease(path, DataTable_source, optionListSource, True) = True Then
					TB_SourceIndicatorLight.BackColor = Color.LightGreen
					B_CopyReleaseFiles.Enabled = True
				Else
					TB_SourceIndicatorLight.BackColor = Color.Red
					B_CopyReleaseFiles.Enabled = False
				End If
			End If
		End If

		DGV_Source.DataSource = Nothing
		DGV_Source.DataSource = DataTable_source
		DGV_Source.ClearSelection()
		DGV_Source.AutoResizeColumns(DataGridViewAutoSizeColumnMode.AllCells)
	End Sub

	Private Sub LoadDestinationGrid()
		CkB_ReadOnly.Checked = False
		DataTable_destination = New DataTable
		DataTable_destination.Columns.Add("Name")
		DataTable_destination.Columns.Add("Date")

		'If this is not our ALPHA location then we need to add in our folders. (+) shows that it is a folder.
		If My.Settings.DestinationFolderPath.ToLower.Contains("alpha") = False Then
			For Each dir As String In Directory.GetDirectories(My.Settings.DestinationFolderPath)
				Dim fileInformation As New DirectoryInfo(dir)
				DataTable_destination.Rows.Add("+ " & fileInformation.Name, fileInformation.LastWriteTime)
				For Each item As FileInfo In fileInformation.GetFiles
					If item.IsReadOnly = True Then
						CkB_ReadOnly.Checked = True
					End If
				Next
			Next
		End If

		'Grab each of the files.
		For Each dir As String In Directory.GetFiles(My.Settings.DestinationFolderPath)
			Dim fileInformation As New FileInfo(dir)
			DataTable_destination.Rows.Add(fileInformation.Name, fileInformation.LastWriteTime)
			If fileInformation.IsReadOnly = True Then
				CkB_ReadOnly.Checked = True
			End If
		Next

		'If this is not our ALPHA location, then we need to preform a series of checks to make sure this location is release ready.
		'Any issues are added at the bottom of the DGV with information as to what is wrong.
		If My.Settings.DestinationFolderPath.ToLower.Contains("alpha") = False Then
			If CheckforRelease(My.Settings.DestinationFolderPath, DataTable_destination, optionListDestination, False) = True Then
				Dim realRelease As Boolean = True
				Dim issues As New List(Of String)

				For Each row In DataTable_destination.Rows
					Dim dr() As DataRow
					dr = DataTable_source.Select("[Name] = '" & row("Name") & "'")
					If dr.Length <> 0 Then
						If row("Name").ToString.Contains("+ Released") = True Then
							Dim source() As String = Directory.GetFiles(My.Settings.SourceFolderPath & "\Released")

							For Each gerber In source
								Dim sourceInformation As New FileInfo(gerber)

								'Check to see if the file exits.
								If File.Exists(My.Settings.DestinationFolderPath & "\Released\" & sourceInformation.Name) Then
									Dim destinationInformation As New FileInfo(My.Settings.DestinationFolderPath & "\" & dr(0)("Name").ToString.Substring(2) & "\" & sourceInformation.Name)

									'Check to see if the file is outdated.
									If sourceInformation.LastWriteTime = destinationInformation.LastWriteTime = False Then
										realRelease = False
										issues.Add("Issue " & sourceInformation.Name & " is outdated")
									End If
								Else
									realRelease = False
									issues.Add("Issue " & sourceInformation.Name & " not found")
								End If
							Next
						Else
							If row("Date") = dr(0)("Date") = False Then
								realRelease = False
								issues.Add("Issue " & row("Name") & " is outdated")
							End If
						End If
					Else
						realRelease = False
						issues.Add("Issue " & row("Name") & " is not in the source")
					End If
				Next

				If realRelease = True Then
					TB_DestinationIndicatorLight.BackColor = Color.LightGreen
				Else
					For Each issue In issues
						DataTable_destination.Rows.Add(issue)
					Next
					TB_DestinationIndicatorLight.BackColor = Color.Red
				End If
			Else
				TB_DestinationIndicatorLight.BackColor = Color.Red
			End If
		Else
			TB_DestinationIndicatorLight.BackColor = Color.Black
		End If

		DGV_Destination.DataSource = Nothing
		DGV_Destination.DataSource = DataTable_destination
		DGV_Destination.ClearSelection()
		DGV_Destination.AutoResizeColumns(DataGridViewAutoSizeColumnMode.AllCells)
	End Sub

	Private Sub LoadSourceAlphaGrid()
		DataTable_source = New DataTable
		DataTable_source.Columns.Add("Name")
		DataTable_source.Columns.Add("Date")

		For Each dir As String In Directory.GetFiles(ALPHA_BACKUP)
			Dim fileInformation As New FileInfo(dir)
			DataTable_source.Rows.Add(fileInformation.Name, fileInformation.LastWriteTime)
		Next

		'Change our indicator light so we do not get confused if these location is release ready.
		TB_SourceIndicatorLight.BackColor = Color.Black
		B_CopyReleaseFiles.Enabled = False

		DGV_Source.DataSource = Nothing
		DGV_Source.DataSource = DataTable_source
		DGV_Source.ClearSelection()
		DGV_Source.AutoResizeColumns(DataGridViewAutoSizeColumnMode.AllCells)
	End Sub

	Private Sub CopyFile(ByRef sourceCopy As String, ByRef sourcePaste As String, ByRef fileName As String)
		Try
			'Try to delete the file first.
			File.Delete(sourcePaste & "\" & fileName)
		Catch ex As Exception
			MsgBox(ex.Message)
		End Try
		File.Copy(sourceCopy & "\" & fileName, sourcePaste & "\" & fileName, True)
		Dim fileinformation As New FileInfo(sourceCopy & "\" & fileName)
		File.SetLastWriteTime(sourcePaste & "\" & fileName, fileinformation.LastWriteTime)
	End Sub

	Private Sub CopyDirectory(ByRef source As String, ByRef destination As String, ByRef folderName As String)
		Try
			'Try to delete the directory first.
			Directory.Delete(destination & "\" & folderName, True)
		Catch ex As Exception

		End Try
		Try
			Dim fileinformation As New DirectoryInfo(destination & "\" & folderName)
			FileSystem.CopyDirectory(source & "\" & folderName, destination & "\" & folderName)
		Catch ex As Exception
			MsgBox(ex.Message)
		End Try
	End Sub

	Private Sub B_ALPHAdestination_Click() Handles B_ALPHAdestination.Click
		Dim selectFolder As New OpenFileDialog
		selectFolder.ValidateNames = False
		selectFolder.CheckFileExists = False
		selectFolder.CheckPathExists = True
		selectFolder.InitialDirectory = RELEASE

		'Used so that we can select empty folder locations as long as the user does not change the file name to a blank name.
		selectFolder.FileName = "Press OK"

		If selectFolder.ShowDialog() = DialogResult.OK Then
			My.Settings.DestinationFolderPath = selectFolder.FileName.Substring(0, selectFolder.FileName.LastIndexOf("\"))
			My.Settings.Save()

			TB_DestinationFolderPath.Text = selectFolder.FileName.Substring(0, selectFolder.FileName.LastIndexOf("\"))
			LoadDestinationGrid()
		End If
	End Sub

	Private Sub B_Check_Click() Handles B_Check.Click
		TB_SourceFolderPath.Text = My.Settings.SourceFolderPath
		LoadSourceGrid(My.Settings.SourceFolderPath)

		SearchForPartner()
	End Sub

	Private Sub TP_QB_items_Enter() Handles TP_QB_items.Enter
		fromTP_Compare = False
	End Sub

	Private Sub TP_PCAD_Build_Enter() Handles TP_PCAD_Build.Enter
		fromTP_Compare = False
	End Sub

	Private Sub TP_BOM_compare_Enter() Handles TP_BOM_compare.Enter
		fromTP_Compare = True
	End Sub

	Private Sub TP_Release_Enter() Handles TP_Release.Enter
		If fromTP_Compare = True Then
			My.Settings.SourceFolderPath = My.Settings.BOMFilePath.Substring(0, My.Settings.BOMFilePath.LastIndexOf("\"))
			My.Settings.Save()
		End If

		TB_SourceFolderPath.Text = My.Settings.SourceFolderPath
		TB_DestinationFolderPath.Text = My.Settings.DestinationFolderPath

		If TB_SourceFolderPath.Text.Length <> 0 Then
			LoadSourceGrid(TB_SourceFolderPath.Text)
		End If
		If My.Settings.DestinationFolderPath.Length <> 0 Then
			SearchForPartner()
		End If
		TP_Release.Show()

		DGV_Source.AutoResizeColumns(DataGridViewAutoSizeColumnMode.AllCells)
		DGV_Destination.AutoResizeColumns(DataGridViewAutoSizeColumnMode.AllCells)
	End Sub
#End Region

#Region "Tab 4: PCAD Build"
	Private Sub BuildBoard_Button_Click() Handles BuildBoard_Button.Click
		Cursor = Cursors.WaitCursor

		fromPCADdatabase = True
		fromSearch = False
		fromCompareItems = False
		fromCompareBOM = False
		B_CompareQBItems.Enabled = True
		B_CompareQBBOM.Enabled = True
		B_Compare_ALPHA.Enabled = True

		If LOGDATA = True Then
			Try
				If ChangeCheck(True) = True Then
					BuildBoards()
				End If
			Catch ex As Exception
				UnhandledExceptionMessage(ex)
			End Try
		Else
			If ChangeCheck(True) = True Then
				BuildBoards()
			End If
		End If
		Cursor = Cursors.Default
	End Sub

	Private Sub BuildBoards()
		'Check to see if we are in the middle of an import.
		Dim message As String = ""
		If sqlapi.CheckDirtyBit(message) = True Then
			MsgBox(message)
			Return
		End If

		'Check that the stop level is a positive whole number
		Try
			numberToBuild = CInt(BuildNumber_TextBox.Text)

			If numberToBuild < 0 Then
				MsgBox("Please input a positive whole number for the number to build.")
				Return
			End If
		Catch ex As Exception
			MsgBox("Please input a positive whole number for the number to build.")
			Return
		End Try

		'Fill up our PCAD BOM table.
		Dim cmdString As String = ""

		If fromSearch = True Then
			cmdString = "SELECT Count(*) AS '" & HEADER_QTY_NEEDED & "', " &
									"[" & DB_HEADER_ITEM_PREFIX & "], " &
									"[" & DB_HEADER_ITEM_NUMBER & "], " &
									"[" & DB_HEADER_MPN & "], " &
									"[" & DB_HEADER_VENDOR & "], " &
									"[" & DB_HEADER_PROCESS & "] " &
									"FROM " & TABLE_TEMP_PCADBOM & " " &
									"WHERE [" & DB_HEADER_PROCESS & "] != '" & PROCESS_NOTUSED & "' AND  [" & DB_HEADER_REF_DES & "] NOT LIKE '" & REFERENCE_DESIGNATOR_OPTION & "%' " &
									"GROUP BY [" & DB_HEADER_ITEM_NUMBER & "],[" & DB_HEADER_ITEM_PREFIX & "],[" & DB_HEADER_MPN & "],[" & DB_HEADER_VENDOR & "],[" & DB_HEADER_PROCESS & "]" &
									"ORDER BY [" & DB_HEADER_ITEM_NUMBER & "]"
		Else
			cmdString = "SELECT Count(*) AS '" & HEADER_QTY_NEEDED & "', " &
									"[" & DB_HEADER_ITEM_PREFIX & "], " &
									"[" & DB_HEADER_ITEM_NUMBER & "], " &
									"[" & DB_HEADER_MPN & "], " &
									"[" & DB_HEADER_VENDOR & "], " &
									"[" & DB_HEADER_PROCESS & "] " &
									"FROM " & TABLE_PCADBOM & " " &
									"WHERE [" & DB_HEADER_BOARD_NAME & "] = '" & BoardBuild_ComboBox.Text & "' AND [" & DB_HEADER_PROCESS & "] != '" & PROCESS_NOTUSED & "' AND  [" & DB_HEADER_REF_DES & "] NOT LIKE '" & REFERENCE_DESIGNATOR_OPTION & "%' " &
									"GROUP BY [" & DB_HEADER_ITEM_NUMBER & "],[" & DB_HEADER_ITEM_PREFIX & "],[" & DB_HEADER_MPN & "],[" & DB_HEADER_VENDOR & "],[" & DB_HEADER_PROCESS & "]" &
									"ORDER BY [" & DB_HEADER_ITEM_NUMBER & "]"
		End If

		Dim da = New SqlDataAdapter(cmdString, myConn)
		build_ds = New DataSet()
		da.Fill(build_ds, "PCAD DATA")

		'Fill up our inventory table. 
		da = New SqlDataAdapter("SELECT [" & DB_HEADER_ITEM_NUMBER & "], " &
								"[" & DB_HEADER_ITEM_PREFIX & "], " &
								"[" & DB_HEADER_TYPE & "], " &
								"[" & DB_HEADER_QUANTITY & "] AS '" & HEADER_QTY_AVAIL & "', " &
								"[" & DB_HEADER_COST & "], " &
								"[" & DB_HEADER_VENDOR & "], " &
								"[" & DB_HEADER_MPN & "], " &
								"[" & DB_HEADER_VENDOR2 & "], " &
								"[" & DB_HEADER_MPN2 & "], " &
								"[" & DB_HEADER_VENDOR3 & "], " &
								"[" & DB_HEADER_MPN3 & "], " &
								"[" & DB_HEADER_LEAD_TIME & "], " &
								"[" & DB_HEADER_MIN_ORDER_QTY & "], " &
								"[" & DB_HEADER_QUANTITY & "] AS '" & HEADER_QTY_ORIG & "' FROM " & TABLE_QB_ITEMS, myConn)
		Dim InventoryDS = New DataSet()
		da.Fill(InventoryDS, "Inventory")

		Dim tableResutlts As New DataTable
		tableResutlts.Columns.Add(DB_HEADER_ITEM_PREFIX)
		tableResutlts.Columns.Add(DB_HEADER_ITEM_NUMBER)
		tableResutlts.Columns.Add(DB_HEADER_MPN)
		tableResutlts.Columns.Add(DB_HEADER_VENDOR)
		tableResutlts.Columns.Add(DB_HEADER_PROCESS)
		tableResutlts.Columns.Add(HEADER_QTY_AVAIL)
		tableResutlts.Columns.Add(HEADER_QTY_NEEDED, GetType(Integer))
		tableResutlts.Columns.Add(HEADER_REMAINDER, GetType(Integer))

		For Each dsrow As DataRow In build_ds.Tables(0).Rows
			Dim drs() As DataRow = InventoryDS.Tables(0).Select("[" & DB_HEADER_ITEM_NUMBER & "] = '" & dsrow(DB_HEADER_ITEM_NUMBER) & "'")

			If drs.Length <> 0 Then
				Dim remainder As Integer = drs(0)(HEADER_QTY_ORIG) - (dsrow(HEADER_QTY_NEEDED) * numberToBuild)

				If ShowAll_CheckBox.Checked = True Then
					tableResutlts.Rows.Add(dsrow(DB_HEADER_ITEM_PREFIX), dsrow(DB_HEADER_ITEM_NUMBER), dsrow(DB_HEADER_MPN), dsrow(DB_HEADER_VENDOR), dsrow(DB_HEADER_PROCESS), CInt(drs(0)(HEADER_QTY_ORIG)), (dsrow(HEADER_QTY_NEEDED) * numberToBuild), remainder)
				ElseIf 0 > remainder Then
					tableResutlts.Rows.Add(dsrow(DB_HEADER_ITEM_PREFIX), dsrow(DB_HEADER_ITEM_NUMBER), dsrow(DB_HEADER_MPN), dsrow(DB_HEADER_VENDOR), dsrow(DB_HEADER_PROCESS), CInt(drs(0)(HEADER_QTY_ORIG)), (dsrow(HEADER_QTY_NEEDED) * numberToBuild), remainder)
				End If
			Else
				'We are not in the database
				Dim remainder As Integer = 0 - (dsrow(HEADER_QTY_NEEDED) * numberToBuild)

				'Add the row anyways because we will have a negitice quantity.
				tableResutlts.Rows.Add(dsrow(DB_HEADER_ITEM_PREFIX), dsrow(DB_HEADER_ITEM_NUMBER), dsrow(DB_HEADER_MPN), dsrow(DB_HEADER_VENDOR), dsrow(DB_HEADER_PROCESS), NOT_IN_DATABASE, (dsrow(HEADER_QTY_NEEDED) * numberToBuild), remainder)
			End If
		Next

		build_ds = New DataSet
		build_ds.Tables.Add(tableResutlts)

		Build_DGV.DataSource = Nothing
		Build_DGV.DataSource = build_ds.Tables(0)
		Build_DGV.AutoResizeColumns(DataGridViewAutoSizeColumnMode.AllCells)

		FormatDGV()
		Build_DGV.Columns(1).Frozen = True
	End Sub

	Private Sub BuildBoardSearch_B_Click() Handles BuildBoardSearch_B.Click
		Cursor = Cursors.WaitCursor
		If LOGDATA = True Then
			Try
				BuildSearchBOM()

				BuildBoards()
			Catch ex As Exception
				UnhandledExceptionMessage(ex)
			End Try
		Else
			BuildSearchBOM()

			BuildBoards()
		End If
		Cursor = Cursors.Default
	End Sub

	Private Sub BuildSearchBOM()
		Dim selectFile As New OpenFileDialog()
		selectFile.InitialDirectory = My.Settings.BOMFilePath
		selectFile.Filter = "All files (*.*)|*.*|All files (*.*)|*.*"
		If selectFile.ShowDialog() = DialogResult.OK Then
			My.Settings.BOMFilePath = selectFile.FileName
			My.Settings.Save()

			Dim originalName As String = Path.GetFileName(selectFile.FileName)
			Dim fileNameParsed() As String = originalName.Split(".")

			If fileNameParsed.Length < 4 Then
				MsgBox("The file you have selected does not parse into 5 parts (boardname).(Rev#).(#).(Option).(bom)" & vbNewLine & "Please check the name of your file for correct naming conventions.")
				Return
			End If

			Dim fileName As String = fileNameParsed(INDEX_BOARD) & "." & fileNameParsed(INDEX_REVISION1) & "." & fileNameParsed(INDEX_REVISION2) & "." & fileNameParsed(INDEX_OPTION) & "."

			ParseBOMFile(My.Settings.BOMFilePath, False)

			TB_FilePath.Text = My.Settings.BOMFilePath

			L_Board.Text = "File: " & fileName
		End If

		DGV_PCAD_BOM.Width = TP_BOM_compare.Width - 15
		DGV_PCAD_BOM.AutoResizeColumns(DataGridViewAutoSizeColumnMode.AllCells)

		DGV_QB_BOM.Visible = False

		fromPCADdatabase = False
		fromSearch = True
		fromCompareItems = False
		fromCompareBOM = False
		B_CompareQBItems.Enabled = True
		B_CompareQBBOM.Enabled = True
		B_Compare_ALPHA.Enabled = True
	End Sub

	Private Sub PopulateSearchBuildDataTable(ByRef filePath As String)
		PCAD_BOM_myCmd.CommandText = "SELECT [" & DB_HEADER_REF_DES & "], [" & DB_HEADER_DESCRIPTION & "], [" & DB_HEADER_MPN & "], COALESCE(NULLIF([" & DB_HEADER_ITEM_PREFIX & "] + ':' + [" & DB_HEADER_ITEM_NUMBER & "],':'), [" & DB_HEADER_ITEM_NUMBER & "]) AS '" & DB_HEADER_ITEM_NUMBER & "', [" & DB_HEADER_VENDOR & "], [" & DB_HEADER_PROCESS & "], [" & DB_HEADER_OPTION & "], [" & DB_HEADER_SWAP & "], [" & DB_HEADER_ERRORS & "] FROM " & TABLE_TEMP_PCADBOM & " ORDER BY [" & DB_HEADER_REF_DES & "]"
		PCAD_BOM_da = New SqlDataAdapter(PCAD_BOM_myCmd)
		PCAD_BOM_ds = New DataSet()

		PCAD_BOM_da.Fill(PCAD_BOM_ds, 0, 500, "PCAD")

		DGV_PCAD_BOM.DataSource = Nothing
		DGV_PCAD_BOM.DataSource = PCAD_BOM_ds.Tables("PCAD")
	End Sub

	Private Sub FormatDGV()
		'Go through the DGV and hilight the different alerts that we want the user to be able to see right away.
		For index = 0 To Build_DGV.Rows.Count - 1

			'Look for a Nigitive remainder item.
			If 0 > Build_DGV.Rows(index).Cells(HEADER_REMAINDER).Value = True Then
				Build_DGV.Rows(index).Cells(DB_HEADER_ITEM_NUMBER).Style.BackColor = OUT_COLOR
			End If

			'Look for a "NOT IN DATABASE" item.
			If Build_DGV.Rows(index).Cells(HEADER_QTY_AVAIL).Value.contains(NOT_IN_DATABASE) = True Then
				Build_DGV.Rows(index).Cells(DB_HEADER_ITEM_NUMBER).Style.BackColor = DATABASE_COLOR
			End If
		Next
	End Sub

	Private Sub Build_DGV_RowPostPaint(ByVal sender As Object, ByVal e As DataGridViewRowPostPaintEventArgs) Handles Build_DGV.RowPostPaint
		'Go through each row of the DGV and add the row number to the row header.
		Using b As SolidBrush = New SolidBrush(Build_DGV.RowHeadersDefaultCellStyle.ForeColor)
			e.Graphics.DrawString(e.RowIndex + 1, Build_DGV.DefaultCellStyle.Font, b, e.RowBounds.Location.X + 10, e.RowBounds.Location.Y + 4)
		End Using
	End Sub

	Private Sub BuildBoardReload_B_Click() Handles BuildBoardReload_B.Click
		Cursor = Cursors.WaitCursor
		If LOGDATA = True Then
			Try
				ReloadBuildSearch()

				BuildBoards()
			Catch ex As Exception
				UnhandledExceptionMessage(ex)
			End Try
		Else
			ReloadBuildSearch()

			BuildBoards()
		End If
		Cursor = Cursors.Default
	End Sub

	Private Sub ReloadBuildSearch()
		Dim originalName As String = Path.GetFileName(My.Settings.BOMFilePath.Substring(My.Settings.BOMFilePath.LastIndexOf("\")))
		Dim fileNameParsed() As String = originalName.Split(".")
		If fileNameParsed.Length < 4 Then
			MsgBox("The file you have selected does not parse into 5 parts (boardname).(Rev#).(#).(Option).(bom)" & vbNewLine & "Please check the name of your file for correct naming conventions.")
			Return
		End If
		Dim fileName As String = fileNameParsed(INDEX_BOARD) & "." & fileNameParsed(INDEX_REVISION1) & "." & fileNameParsed(INDEX_REVISION2) & "." & fileNameParsed(INDEX_OPTION) & "."

		L_Board.Text = "File: " & fileName

		myCmd.CommandText = "DELETE FROM " & TABLE_TEMP_PCADBOM
		myCmd.ExecuteNonQuery()

		ParseBOMFile(My.Settings.BOMFilePath, True)

		TB_FilePath.Text = My.Settings.BOMFilePath

		DGV_PCAD_BOM.Width = TP_BOM_compare.Width - 15
		DGV_PCAD_BOM.AutoResizeColumns(DataGridViewAutoSizeColumnMode.AllCells)

		DGV_QB_BOM.Visible = False

		fromPCADdatabase = False
		fromSearch = True
		fromCompareItems = False
		fromCompareBOM = False
		B_CompareQBItems.Enabled = True
		B_CompareQBBOM.Enabled = True
		B_Compare_ALPHA.Enabled = True
	End Sub


	Private Sub GenerateALPHA_Button_Click() Handles GenerateALPHA_Button.Click
		Cursor = Cursors.WaitCursor
		If LOGDATA = True Then
			Try
				GenerateALPHA()
			Catch ex As Exception
				UnhandledExceptionMessage(ex)
			End Try
		Else
			GenerateALPHA()
		End If
		Cursor = Cursors.Default
	End Sub

	Private Sub GenerateALPHA()
		Dim message As String = ""
		If sqlapi.CheckDirtyBit(message) = True Then
			MsgBox(message)
			Return
		End If

		Dim boardName As String = ""
		Dim pnpFilePath As String = ""

		' First get the pnp file that will be parsed.
		Dim selectFile As New OpenFileDialog()
		selectFile.InitialDirectory = My.Settings.BOMFilePath
		selectFile.Filter = "All files (*.*)|*.*|All files (*.*)|*.*"
		If selectFile.ShowDialog() = DialogResult.OK Then

			pnpFilePath = selectFile.FileName
			Dim fileNameParsed() As String = Path.GetFileName(selectFile.FileName).Split(".")

			If fileNameParsed.Length < 6 Then
				MsgBox("The file you have selected does not parse into 6 parts (boardname).(Rev#).(#).(Option).(pnp.csv)" & vbNewLine & "Please check the name of your file for correct naming conventions.")
				Return
			End If

			Dim enterOption As New MessageBoxOption()

			If enterOption.ShowDialog() = DialogResult.Cancel Then
				Return
			End If

			boardName = fileNameParsed(INDEX_BOARD) & "." & fileNameParsed(INDEX_REVISION1) & "." & fileNameParsed(INDEX_REVISION2) & "." & enterOption.TB_Options.Text.ToUpper & "."

			CreatePNPList(pnpFilePath, boardName)
		End If
	End Sub

	Private Function RotationX(ByRef bR As Integer, ByRef posX As Integer, ByRef aX As Integer,
								ByRef posY As Integer, ByRef aY As Integer) As Integer
		Dim answer As Integer

		'Depending on our rotation of the board, we need to use the correct formula to change where the new X position will be.
		'aX/aY are the offsets from the fiducial mark.
		Select Case bR
			Case 0
				answer = (posX - aX) * MICROMETTER_CONVERTER
			Case 90
				answer = (-posY + aY) * MICROMETTER_CONVERTER
			Case 180
				answer = (-posX + aX) * MICROMETTER_CONVERTER
			Case 270
				answer = (posY - aY) * MICROMETTER_CONVERTER
		End Select
		Return answer
	End Function

	Private Function RotationY(ByRef bR As Integer, ByRef posX As Integer, ByRef aX As Integer,
								ByRef posY As Integer, ByRef aY As Integer) As Integer
		Dim answer As Integer

		'Depending on our rotation of the board, we need to use the correct formula to change where the new Y position will be.
		'aX/aY are the offsets from the fiducial mark.
		Select Case bR
			Case 0
				answer = (posY - aY) * MICROMETTER_CONVERTER
			Case 90
				answer = (posX - aX) * MICROMETTER_CONVERTER
			Case 180
				answer = (-posY + aY) * MICROMETTER_CONVERTER
			Case 270
				answer = (-posX + aX) * MICROMETTER_CONVERTER
		End Select
		Return answer
	End Function

	Private Sub MirrorOverYAxis(ByRef posX As Integer)
		posX = 0 - posX
	End Sub

	Private Sub CreatePNPList(ByRef pnpPath As String, ByRef boardName As String)
		If ParsePNPFile(pnpPath, boardName) = False Then
			Return
		End If

		Dim myReader As SqlDataReader = Nothing
		Dim reportList As New List(Of String)

		'Create the header for the ALPHA file.
		reportList.Add("# *** PCBS ***")
		reportList.Add("F1 " & boardName)
		reportList.Add("F21 All_Tools")

		'X/Y offset (also known as X/Y coordinents for fiducial mark 1).
		Dim AX As Integer = fu1X
		Dim AXMirror As Integer = 0
		Dim AY As Integer = fu1Y

		'Variables used to set what the new X/Y coordinent will be after calculating offset.
		Dim newX As Integer = 0
		Dim newY As Integer = 0

		newX = RotationX(boardRotation, fu1X, AX, fu1Y, AY)
		newY = RotationY(boardRotation, fu1X, AX, fu1Y, AY)

		'These should equal 0.
		Dim newfu1X = newX
		Dim newfu1Y = newY

		newX = RotationX(boardRotation, fu2X, AX, fu2Y, AY)
		newY = RotationY(boardRotation, fu2X, AX, fu2Y, AY)

		Dim newfu2X = newX
		Dim newfu2Y = newY

		newX = RotationX(boardRotation, fu3X, AX, fu3Y, AY)
		newY = RotationY(boardRotation, fu3X, AX, fu3Y, AY)

		Dim newfu3X = newX
		Dim newfu3Y = newY

		If boardMirror <> 0 Then
			AXMirror = newfu3X
			MirrorOverYAxis(newfu1X)
			newfu1X = newfu1X + AXMirror
			MirrorOverYAxis(newfu2X)
			newfu2X = newfu2X + AXMirror
			MirrorOverYAxis(newfu3X)
			newfu3X = newfu3X + AXMirror
		End If

		'Add the fiducial marks to the ALPHA file.
		reportList.Add("F3 " & newfu1X & " " & newfu1Y & " " & fiducialName)
		reportList.Add("F3 " & newfu2X & " " & newfu2Y & " " & fiducialName)
		reportList.Add("F3 " & newfu3X & " " & newfu3Y & " " & fiducialName)

		'Create and add our query into our data set for comparison.
		Dim myCmd As New SqlCommand("SELECT * FROM " & TABLE_TEMP_PNP & " ORDER BY [" & DB_HEADER_REF_DES & "]", myConn)
		Dim Temp_PNP As New DataTable()
		Temp_PNP.Load(myCmd.ExecuteReader)

		'Get new locations for all of the parts.
		For Each dsRow As DataRow In Temp_PNP.Rows
			newX = RotationX(boardRotation, dsRow(DB_HEADER_POS_X), AX, dsRow(DB_HEADER_POS_Y), AY)
			newY = RotationY(boardRotation, dsRow(DB_HEADER_POS_X), AX, dsRow(DB_HEADER_POS_Y), AY)

			If boardMirror <> 0 Then
				MirrorOverYAxis(newX)
				newX = newX + AXMirror
			End If

			'Add the new information to the ALPHA file.
			'F8 is where we put the X location, Y location, rotation, group, mount-skip, dispense-skip, component
			'F9 is wehre we put the reference designator.
			'Hard coded rotation '0' to force manual check of the pnp with each part.
			reportList.Add("F8 " & newX & " " & newY & " 0 0 N N " & dsRow(DB_HEADER_ITEM_NUMBER))
			reportList.Add("F9 " & dsRow(DB_HEADER_REF_DES))
		Next

		Dim report As New GenerateReport()
		GenerateALPHAfile(reportList, boardName)
	End Sub

	Public Sub GenerateALPHAfile(ByRef list As List(Of String), ByRef boardName As String)
		Try
			Dim TempFile As New StreamWriter("\\Server1\Shares\Production\AlphaBackup\" & boardName & ".gen", False)
			For Each item In list
				TempFile.WriteLine(item)
			Next

			TempFile.Close()
		Catch ex As Exception
			MsgBox(ex.Message)
		End Try
	End Sub

	Private Function ParsePNPFile(ByRef path As String, ByRef boardName As String) As Boolean
		'Clear the Database for the new PNP file that we are going to create.
		Dim myCmd As New SqlCommand("DELETE FROM " & TABLE_TEMP_PNP, myConn)
		myCmd.ExecuteNonQuery()
		'Dim originalName As String = board
		Dim fileNameParsed() As String = boardName.Split(".")

		'Indexs
		Dim INDEX_refdes As Integer = -1
		Dim INDEX_stockNumber As Integer = -1
		Dim INDEX_PosX As Integer = -1
		Dim INDEX_PosY As Integer = -1
		Dim INDEX_Rotation As Integer = -1
		Dim INDEX_Process As Integer = -1
		Dim INDEX_Value As Integer = -1

		'Optional
		Dim INDEX_options As Integer = -1
		Dim INDEX_swap As Integer = -1

		Dim fu1_check As Integer = -1
		Dim fu2_check As Integer = -1
		Dim fu3_check As Integer = -1
		Dim br_check As Integer = -1
		Dim fn_check As Integer = -1
		Dim mi_check As Integer = -1

		Dim lineNo = 0

		'Start our transaction. Must assign both transaction object and connection to the command object for a pending local transaction.
		Dim transaction As SqlTransaction = Nothing
		transaction = myConn.BeginTransaction("Temp Transaction")
		myCmd.Connection = myConn
		myCmd.Transaction = transaction

		Try
			Dim foundIssue As Boolean = False
			Using myParser As New TextFieldParser(path)
				myParser.TextFieldType = FieldType.Delimited
				myParser.SetDelimiters(",")
				Dim currentRow As String()

				'First three rows are the header. We do not need any of this information.
				currentRow = myParser.ReadFields()
				lineNo += 1
				currentRow = myParser.ReadFields()
				lineNo += 1
				currentRow = myParser.ReadFields()
				lineNo += 1
				Dim index As Integer = 0
				Dim missingFields As String = ""
				Dim fieldErrors As Boolean = False

				'Parse the header row to grab Indexs. They can be generated in any order.
				For Each header In currentRow
					Select Case header.ToLower
						Case "refdes"
							INDEX_refdes = index
						Case "stock number"
							INDEX_stockNumber = index
						Case "locationx"
							INDEX_PosX = index
						Case "locationy"
							INDEX_PosY = index
						Case "rotation"
							INDEX_Rotation = index
						Case "option"
							INDEX_options = index
						Case "process"
							INDEX_Process = index
						Case "swap"
							INDEX_swap = index
						Case "value"
							INDEX_Value = index
					End Select
					index += 1
				Next

				'Check to see if we are missing the important fields.
				If INDEX_refdes = -1 Then
					fieldErrors = True
					missingFields = missingFields & "refdes |"
				End If
				If INDEX_stockNumber = -1 Then
					fieldErrors = True
					missingFields = missingFields & " stockNumber |"
				End If
				If INDEX_PosX = -1 Then
					fieldErrors = True
					missingFields = missingFields & " locationX |"
				End If
				If INDEX_PosY = -1 Then
					fieldErrors = True
					missingFields = missingFields & " locationY |"
				End If
				If INDEX_Rotation = -1 Then
					fieldErrors = True
					missingFields = missingFields & " Rotation |"
				End If
				If INDEX_Process = -1 Then
					fieldErrors = True
					missingFields = missingFields & " Process |"
				End If
				If INDEX_Value = -1 Then
					fieldErrors = True
					missingFields = missingFields & " Value"
				End If

				If fieldErrors = True Then
					MsgBox("PNP File is missing the following Fields: " & missingFields)
					sqlapi.RollBack(transaction, errorMessage:=New List(Of String))
					Return False
				End If

				While Not myParser.EndOfData
					Dim referenceDesignator As String = ""
					Dim itemPrefix As String = ""
					Dim itemNumber As String = ""
					Dim posX As String = ""
					Dim posY As String = ""
					Dim rotation As String = ""
					Dim process As String = ""
					'Optional
					Dim optionValue As String = ""
					currentRow = myParser.ReadFields()
					lineNo += 1

					If currentRow(0).Length = 0 Then
						Continue While
					End If

					'Check to see if we have the option field or not.
					If INDEX_options <> -1 Then
						Dim include As Boolean = False
						optionValue = currentRow(INDEX_options)

						'Check to see if we have an option.
						If optionValue.Length <> 0 Then
							For index = 0 To optionValue.Length - 1

								'Check each letter of the option feild to see if the file calls for it.
								If fileNameParsed(INDEX_OPTION).Contains(optionValue(index)) = True Then
									include = True
									Exit For
								End If
							Next
							If include = False Then
								Continue While
							End If
						End If
					End If

					'- - - Parse Reference Designator - - -

					If INDEX_swap <> -1 Then
						'Check to see if we have a swap.
						If currentRow(INDEX_swap).Length <> 0 Then
							referenceDesignator = currentRow(INDEX_swap)
						Else
							referenceDesignator = currentRow(INDEX_refdes)
						End If
					Else
						referenceDesignator = currentRow(INDEX_refdes)
					End If

					'Check to see if we are wroking with our FUs.
					Try
						'The 'FU' Reference Designator is very important. This is where we are storing all of the information
						'	that deals with our fiducial marks on the board. 
						'	We should not have any value with the exception of the FU1 where the board roation and fiducial name are found.
						'	Each should have an x,y coordinate.
						Select Case referenceDesignator
							Case "FU1"
								fu1_check = 0
								fu1X = currentRow(INDEX_PosX)
								fu1Y = currentRow(INDEX_PosY)
								If currentRow(INDEX_Value).Contains("|") = True Then
									Dim parsed() As String = currentRow(INDEX_Value).Split("|")
									For Each item In parsed
										Select Case item.Substring(0, 2).ToLower
											Case "br"
												br_check = 0
												boardRotation = item.Substring(2)
											Case "fn"
												fn_check = 0
												fiducialName = item.Substring(2)
											Case "mi"
												boardMirror = 1
											Case Else
												MsgBox("FU1 '|' format")
												myParser.Close()
												transaction.Rollback()
												Return False
										End Select
									Next
									Continue While
								Else
									MsgBox("FU1 '|' format")
									myParser.Close()
									transaction.Rollback()
									Return False
								End If
							Case "FU2"
								fu2_check = 0
								fu2X = currentRow(INDEX_PosX)
								fu2Y = currentRow(INDEX_PosY)
								Continue While
							Case "FU3"
								fu3_check = 0
								fu3X = currentRow(INDEX_PosX)
								fu3Y = currentRow(INDEX_PosY)
								Continue While
						End Select
					Catch ex As Exception

					End Try

					'- - - Parse Stock Number - - -

					If INDEX_stockNumber <> -1 Then
						itemNumber = currentRow(INDEX_stockNumber).Substring(currentRow(INDEX_stockNumber).IndexOf(":") + 1)

						'- - - Parse Prefix - - -

						'Check to see if we have a colon.
						If currentRow(INDEX_stockNumber).Contains(":") = True Then
							itemPrefix = currentRow(INDEX_stockNumber).Substring(0, currentRow(INDEX_stockNumber).IndexOf(":"))
						End If
					End If

					'- - - Parse X position - - -

					If INDEX_PosX <> -1 Then
						posX = currentRow(INDEX_PosX).Substring(0, currentRow(INDEX_PosX).IndexOf("."))
					End If

					'- - - Parse Y position - - -

					If INDEX_PosY <> -1 Then
						posY = currentRow(INDEX_PosY).Substring(0, currentRow(INDEX_PosY).IndexOf("."))
					End If

					'- - - Parse Rotation - - -

					If INDEX_Rotation <> -1 Then
						rotation = currentRow(INDEX_Rotation).Substring(0, currentRow(INDEX_Rotation).IndexOf("."))
					End If

					'- - - Parse Process - - -

					If INDEX_Process <> -1 Then
						process = currentRow(INDEX_Process)

						'Check to makes sure that we have only 'SMT' process'.
						If String.Compare(process, PROCESS_SMT, True) <> 0 Then
							If String.Compare(process, PROCESS_NOTUSED, True) <> 0 Then
								MsgBox(referenceDesignator & " Process is " & process)
								myParser.Close()
								transaction.Rollback()
								Return False
							Else
								Continue While
							End If
						End If
					End If

					myCmd.CommandText = "INSERT INTO " & TABLE_TEMP_PNP & " ([" & DB_HEADER_REF_DES & "], [" & DB_HEADER_ITEM_PREFIX & "], [" & DB_HEADER_ITEM_NUMBER & "], [" & DB_HEADER_POS_X & "], [" & DB_HEADER_POS_Y & "], [" & DB_HEADER_ROTATION & "], [" & DB_HEADER_PROCESS & "]) " &
										"VALUES('" & referenceDesignator & "','" & itemPrefix & "','" & itemNumber & "','" & posX & "','" & posY & "','" & rotation & "','" & process & "')"
					myCmd.ExecuteNonQuery()

				End While
			End Using

			'Check to see if we have found each of our Fiducial information parts
			'Without them we cannot create an alpha file.
			If br_check = -1 Or fn_check = -1 Or fu1_check = -1 Or fu2_check = -1 Or fu3_check = -1 Then
				Dim infostring As String = "Fiducial Information missing:"
				If br_check = -1 Then
					infostring = infostring & " BR"
				End If
				If fn_check = -1 Then
					infostring = infostring & " FN"
				End If
				If fu1_check = -1 Then
					infostring = infostring & " FU1"
				End If
				If fu2_check = -1 Then
					infostring = infostring & " FU2"
				End If
				If fu3_check = -1 Then
					infostring = infostring & " FU3"
				End If

				MsgBox(infostring)
				transaction.Rollback()
				Return False
			End If

			transaction.Commit()
		Catch ex As Exception
			If Not transaction Is Nothing Then
				sqlapi.RollBack(transaction, errorMessage:=New List(Of String))
				MsgBox(lineNo & ": " & ex.Message)
				Return False
			End If
		End Try
		Return True
	End Function

#End Region

End Class