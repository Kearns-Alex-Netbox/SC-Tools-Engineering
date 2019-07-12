'-----------------------------------------------------------------------------------------------------------------------------------------
' Module: Settings.vb
'
' Description: Settings that are different from user to user. These settings are saved and used throughout the program.
'
' Settings:
'	Project Directory = The folder directory that contains all of the release directories for the database.
'	Database File = The database file that QB uses. Used for checking if we need to preform an update
'	Remind Me = Enable or disable the timer that is used to check the database file for changes.
'-----------------------------------------------------------------------------------------------------------------------------------------
Imports System.Data.Odbc

Public Class Settings

	Private Sub Settings_Load() Handles MyBase.Load
		CenterToParent()
		L_Version.Text = "V:" & Application.ProductVersion & "  " & DATABASE
		TB_ProjectDirectory.Text = My.Settings.ReleaseLocation
		Remind_CheckBox.Checked = My.Settings.RemindMe
		DatabaseFile_TextBox.Text = My.Settings.DatabasePath
		RevisionExcel_TextBox.Text = My.Settings.RevisionFile
	End Sub

	Private Sub BrowseReleaseLocation_Button_Click() Handles BrowseReleaseLocation_Button.Click
		OpenLocation(TB_ProjectDirectory.Text, "Select Project Directory", TB_ProjectDirectory.Text, "foldersOnly|*.none", True)
	End Sub

	Private Sub BrowseDatabaseFile_Button_Click() Handles BrowseDatabaseFile_Button.Click
		OpenLocation(DatabaseFile_TextBox.Text, "Select Database File", DatabaseFile_TextBox.Text, "QuickBooks|*.qbw", False)
	End Sub

	Private Sub RevisionBrowse_Button_Click(sender As Object, e As EventArgs) Handles RevisionBrowse_Button.Click
		OpenLocation(RevisionExcel_TextBox.Text, "Select Revision File", RevisionExcel_TextBox.Text, "Excel|*.xlsx", False)
	End Sub

	Private Sub B_Save_Click() Handles B_Save.Click
		If TB_ProjectDirectory.Text.Length = 0 Then
			MsgBox("Please enter a Project Directory.")
			Return
		End If
		If DatabaseFile_TextBox.Text.Length = 0 Then
			MsgBox("Please enter a path for the database file.")
			Return
		End If

		My.Settings.ReleaseLocation = TB_ProjectDirectory.Text
		My.Settings.DatabasePath = DatabaseFile_TextBox.Text
		My.Settings.RemindMe = Remind_CheckBox.Checked
		My.Settings.RevisionFile = RevisionExcel_TextBox.Text
		My.Settings.Save()

		Close()
	End Sub

	Private Sub B_Cancel_Click() Handles B_Cancel.Click
		Close()
	End Sub

	Private Sub TestConnection_Button_Click() Handles TestConnection_Button.Click
		Try
			_cn.Open()
			MessageBox.Show("Connection Opened")
			_cn.Close()
			MessageBox.Show("Connection Closed")
		Catch ex As OdbcException
			MessageBox.Show(ex.Errors.Item(0).Message)
		End Try
	End Sub

	Private Sub OpenLocation(ByVal root As String, ByVal windowTitle As String, ByRef finalLocation As String, ByVal filter As String, ByVal isFolder As Boolean)
		' method: https://stackoverflow.com/questions/32370524/setting-root-folder-for-folderbrowser
		Using obj As New OpenFileDialog
			obj.Filter = filter
			obj.CheckFileExists = False
			obj.CheckPathExists = False
			obj.InitialDirectory = root
			obj.CustomPlaces.Add("\\Server1")
			obj.CustomPlaces.Add("C:")
			obj.Title = windowTitle

			If isFolder Then
				obj.FileName = "OpenFldrPath"
			End If

			If obj.ShowDialog = Windows.Forms.DialogResult.OK Then
				If isFolder Then
					finalLocation = IO.Directory.GetParent(obj.FileName).FullName
				Else
					finalLocation = obj.FileName
				End If
			End If
		End Using
	End Sub

End Class