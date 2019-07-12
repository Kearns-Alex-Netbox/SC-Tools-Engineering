'-----------------------------------------------------------------------------------------------------------------------------------------
' Module: MenuMain.vb
'
' Description: Main Menu of the whole program. Side panel with buttons that open up forms inside the larger panel area. Top panel has 
'		special buttons that control all of the open forms. The last Import and ALPHA Import date and times are also displayed here as
'		well. Only one form may be open at a time with the exception of the Master Control window.
'
' Buttons:
'	Cascade = Takes all of the open forms and Cascades them from the top left cornor in order of last visited.
'	Tile = Takes all of the open forms and tiles them side by side to a more even size so that every form is on the screen.
'	Minimize = Minimizes all of the open forms.
'	Close = Closes all of the open forms.
'
' Special Keys: These key combinations can be used on any form. DO NOT USE THEM AGAIN ON INDIVIDUAL FORMS.
'   control-S = Opens another search window
'	control-P = Opens the process window
'	1 = ONLY APPLIES to user name [akearns]. Special debugging key that will open up what ever form is programmed to it. Developing forms
'		are assigned here before given a button.
'	2 = ONLY APPLIES to user name [akearns]. Special debugging key that will open up what ever form is programmed to it. Developing forms
'		are assigned here before given a button.
'-----------------------------------------------------------------------------------------------------------------------------------------
Imports System.Data.SqlClient

Public Class MenuMain

	Private Sub MenuMain_Load() Handles MyBase.Load
		CenterToParent()
		My.Settings.RemindMe = True
		My.Settings.Save()

		KeyPreview = True
		IsMdiContainer = True

		Dim myCmd As New SqlCommand("IF EXISTS(SELECT [" & DB_HEADER_NAME & "] FROM " & TABLE_UTILITIES & " WHERE [" & DB_HEADER_NAME & "] = 'LastUpdate') SELECT 1 ELSE SELECT 0", myConn)
		Dim result As Integer = myCmd.ExecuteScalar

		If result = 0 Then
			myCmd = New SqlCommand("INSERT INTO " & TABLE_UTILITIES & " ([" & DB_HEADER_NAME & "], [" & DB_HEADER_VALUE & "]) VALUES('LastUpdate','')", myConn)
			myCmd.ExecuteNonQuery()
		End If

		myCmd.CommandText = "IF EXISTS(SELECT [" & DB_HEADER_NAME & "] FROM " & TABLE_UTILITIES & " WHERE [" & DB_HEADER_NAME & "] = 'LastALPHAUpdate') SELECT 1 ELSE SELECT 0"
		result = myCmd.ExecuteScalar

		If result = 0 Then
			myCmd = New SqlCommand("INSERT INTO " & TABLE_UTILITIES & " ([" & DB_HEADER_NAME & "], [" & DB_HEADER_VALUE & "]) VALUES('LastALPHAUpdate','')", myConn)
			myCmd.ExecuteNonQuery()
		End If

		myCmd = New SqlCommand("SELECT [" & DB_HEADER_VALUE & "] FROM " & TABLE_UTILITIES & " WHERE [" & DB_HEADER_NAME & "] = 'LastUpdate'", myConn)
		L_LastImport.Text = myCmd.ExecuteScalar

		myCmd = New SqlCommand("SELECT [" & DB_HEADER_VALUE & "] FROM " & TABLE_UTILITIES & " WHERE [" & DB_HEADER_NAME & "] = 'LastALPHAUpdate'", myConn)
		L_LastALPHAImport.Text = myCmd.ExecuteScalar

		Dim ctl As Control
		Dim ctlMDI As MdiClient

		' Loop through all of the form's controls looking for the control of type MdiClient.
		For Each ctl In Controls
			Try
				' Attempt to cast the control to type MdiClient.
				ctlMDI = CType(ctl, MdiClient)

				' Set the BackColor of the MdiClient control.
				ctlMDI.BackColor = BackColor
			Catch exc As InvalidCastException
				' Catch and ignore the error if casting failed.
			End Try
		Next

		HasDatabaseChanged()
	End Sub

	Private Sub ImportData_Button_Click(sender As Object, e As EventArgs) Handles ImportData_Button.Click
		If My.Settings.ReleaseLocation.Length = 0 Then
			MsgBox("Project Directory not set up yet. Taking you to Settings.")
			Dim DoSettings As New Settings()
			DoSettings.ShowDialog()
			Return
		End If
		OpenForm(New ImportData(Me))
	End Sub

	Private Sub B_CompareRevisions_Click() Handles B_CompareRevisions.Click
		OpenForm(CompareRevisions)
	End Sub

	Private Sub B_PartUsage_Click() Handles B_PartUsage.Click
		OpenForm(PartUsage)
	End Sub

	Private Sub B_BoardOption_Click() Handles B_BoardOption.Click
		OpenForm(BoardOptions)
	End Sub

	Private Sub B_FindStockNumber_Click() Handles B_FindStockNumber.Click
		OpenForm(FindItemNumber)
	End Sub

	Private Sub B_MasterWindow_Click() Handles B_MasterWindow.Click
		Dim message As String = ""
		If sqlapi.CheckDirtyBit(message) = True Then
			MsgBox(message & vbNewLine & "You will have to open the window once the dirty bit is clear to get all items.")
			Return
		End If
		If ChangeCheck(True) = True Then
			Dim doMasterControl As New MasterControl
			doMasterControl.Show()
		End If
	End Sub

	Private Sub B_Settings_Click() Handles B_Settings.Click
		Dim Settings As New Settings()
		Settings.ShowDialog()
	End Sub

#Region "Top Panel Buttons"
	Private Sub B_Cascade_Click() Handles B_Cascade.Click
		For Each form As Form In MdiChildren
			form.WindowState = FormWindowState.Normal
		Next
		LayoutMdi(MdiLayout.Cascade)
	End Sub

	Private Sub B_Tile_Click() Handles B_Tile.Click
		For Each form As Form In MdiChildren
			form.WindowState = FormWindowState.Normal
		Next
		LayoutMdi(MdiLayout.TileVertical)
	End Sub

	Private Sub B_Minimize_Click() Handles B_Minimize.Click
		For Each form As Form In MdiChildren
			form.WindowState = FormWindowState.Minimized
		Next
	End Sub

	Private Sub B_Close_Click() Handles B_Close.Click
		For Each form As Form In MdiChildren
			form.Close()
		Next
	End Sub
#End Region

	Private Sub B_Exit_Click() Handles B_Exit.Click
		Dim result As String = ""
		sqlapi.CloseDatabase(myConn, result)
		Application.Exit()
	End Sub

	Private Sub MenuMain_FormClosed() Handles Me.FormClosed
		Dim result As String = ""
		sqlapi.CloseDatabase(myConn, result)
		Application.Exit()
	End Sub

	Public Function OpenForm(ByRef thisForm As Form) As Boolean
		'Check to see if we are doing an import only if we are not opening the import data window.
		'Import data is the only place that we can reset the dirty bit.
		If thisForm.Name.Contains("Import") = False Then
			Dim message As String = ""
			If sqlapi.CheckDirtyBit(message) = True Then
				MsgBox(message)
				Return False
			End If
		End If

		Dim indexOfMenu As Integer = 0

		Dim frmCollection = Application.OpenForms
		For i = 0 To frmCollection.Count - 1
			If frmCollection.Item(i).Name = thisForm.Name Then
				frmCollection.Item(i).Activate()
				Return True
			End If
			If frmCollection.Item(i).Name = Name Then
				indexOfMenu = i
			End If
		Next i
		thisForm.StartPosition = FormStartPosition.Manual
		thisForm.Left = 0
		thisForm.Top = 0
		thisForm.MdiParent = frmCollection.Item(indexOfMenu)
		thisForm.Show()
		thisForm.BringToFront()

		Return True
	End Function

	Public Sub MyBase_KeyDown(ByVal sender As Object, ByVal e As KeyEventArgs) Handles MyBase.KeyDown
		'If the key combination is Ctrl-S then we open anthother Master Control window.
		If (e.Control AndAlso (e.KeyCode = Keys.S)) Then
			Dim DoQBItems As New MasterControl
			DoQBItems.Show()
		End If

		'If the key combination is Ctrl-P then we open our PCAD Process window.
		If (e.Control AndAlso (e.KeyCode = Keys.S)) Then
			Dim DoQBItems As New MasterControl
			DoQBItems.Show()
		End If

		'----- TESTING ----- TESTING ----- TESTING ----- TESTING ----- TESTING ----- TESTING ----- TESTING -----'
		If String.Compare(sqlapi._Username, "akearns") = 0 Then
			'If e.KeyCode = Keys.E Then
			'    Throw New Exception("'e' Key Error.")
			'End If
			If e.KeyCode = Keys.D1 Then
				'OpenForm(CriticalBuildPath3a)
			End If
			If e.KeyCode = Keys.D2 Then
				'OpenForm(CriticalBuildPath3a)
			End If
		End If
	End Sub

End Class