'-----------------------------------------------------------------------------------------------------------------------------------------
' Module: MessageBoxQB.vb
'
' Description: Custom message box with a few special tricks. Displays a message that is passed in to the user and has buttons  that allow
'	the user to choose an answer. A small reminder checkbox is also included to give the user the option to see this message again.
'
' Checkboxes:
'	Remind Me Again = Allows the user to turn off the reminder for the rest of the session. This can also be changed inside the settings.
'		This is controled by the last parameter 'isReminder'. If this boolean is passed in as 'false' then this option will not be
'		displayed to the user.
'
' Buttons:
'	One = Always needs to be populated. This button will take the string passed in as 'OneText' and always return 'Yes' when pressed.
'	Two = [Optional] This button will take the string passed in as 'TwoText' and always return 'No' when pressed. This parameter can also
'		have 'Nothing' passed in to result in this button not being used. This will make the button disapear and not clickable.
'	Three = [Optional] This button will take the string passed in as 'ThreeText' and always return 'Ignore' when pressed. This parameter
'		can also have 'Nothing' passed in to result in this button not being used. This will make the button disapear and not clickable.
'
'-----------------------------------------------------------------------------------------------------------------------------------------

Public Class MessageBoxQB

	Dim reminder As Boolean

	''' <summary>
	''' Creates a new instance of the MessageBoxQB. This is a custom message system that allows the messages to be displayed and buttons to
	''' respond to the message if needed.
	''' </summary>
	''' <param name="message">The message that will be displayed to the user.</param>
	''' <param name="OneText">The label that will be on the first button. This is always needed.</param>
	''' <param name="TwoText">The label that will be on the second button. If 'Nothing' is passed then the button will not show.</param>
	''' <param name="ThreeText">The label that will be on the third button. If 'Nothing' is passed then the button will not show.</param>
	''' <param name="isReminder">Boolean that controls if the message will give the option to turn off alerts for future reminders.</param>
	Public Sub New(ByRef message As String, ByRef OneText As String, ByRef TwoText As String, ByRef ThreeText As String, ByRef isReminder As Boolean)

		' This call is required by the designer.
		InitializeComponent()

		' Add any initialization after the InitializeComponent() call.
		Message_RichTextBox.Text = message

		One_Button.Text = OneText

		If TwoText IsNot Nothing Then
			Two_Button.Text = TwoText
		Else
			Two_Button.Text = ""
			Two_Button.Enabled = False
			Two_Button.Visible = False
		End If

		If ThreeText IsNot Nothing Then
			Three_Button.Text = ThreeText
		Else
			Three_Button.Text = ""
			Three_Button.Enabled = False
			Three_Button.Visible = False
		End If

		reminder = isReminder
		If isReminder = False Then
			Remind_CheckBox.Visible = False
		End If
	End Sub

	Private Sub One_Button_Click(sender As Object, e As EventArgs) Handles One_Button.Click
		If reminder = True Then
			My.Settings.RemindMe = Remind_CheckBox.Checked
			My.Settings.Save()
		End If

		DialogResult = DialogResult.Yes
	End Sub

	Private Sub Two_Button_Click(sender As Object, e As EventArgs) Handles Two_Button.Click
		If reminder = True Then
			My.Settings.RemindMe = Remind_CheckBox.Checked
			My.Settings.Save()
		End If

		DialogResult = DialogResult.No
	End Sub

	Private Sub Three_Button_Click(sender As Object, e As EventArgs) Handles Three_Button.Click
		If reminder = True Then
			My.Settings.RemindMe = Remind_CheckBox.Checked
			My.Settings.Save()
		End If

		DialogResult = DialogResult.Ignore
	End Sub

End Class