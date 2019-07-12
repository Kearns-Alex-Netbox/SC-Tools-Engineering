Public Class MessageBoxOption

	Private Sub MessageBoxOption_Load() Handles MyBase.Load
		CenterToParent()
	End Sub

	Private Sub B_OK_Click() Handles B_OK.Click
		DialogResult = DialogResult.OK
		Close()
	End Sub

	Private Sub B_Cancel_Click() Handles B_Cancel.Click
		DialogResult = DialogResult.Cancel
		Close()
	End Sub

End Class