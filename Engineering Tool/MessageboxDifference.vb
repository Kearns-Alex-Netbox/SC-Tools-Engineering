'-----------------------------------------------------------------------------------------------------------------------------------------
' Module: MessageboxDifference.vb
'
' Description: Custom Message window that is used to tell the user what the differences are between two items in two different BOMs.
'
'-----------------------------------------------------------------------------------------------------------------------------------------

Public Class MessageboxDifference

	Dim errorList As List(Of String)

	Public Sub New(ByRef list As List(Of String))
		InitializeComponent()
		errorList = list
	End Sub

	Private Sub CustomMessagebox_Load() Handles MyBase.Load
		CenterToParent()
		For Each item In errorList
			RTB_differenceReport.Text = RTB_differenceReport.Text & item & vbNewLine
		Next
	End Sub

End Class