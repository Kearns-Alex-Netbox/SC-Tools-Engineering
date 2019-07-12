'-----------------------------------------------------------------------------------------------------------------------------------------
' Module: PCADProcess.vb
'
' Description: Shows all of the valid Processes and Fields that are used in the PCAD BOMs.
'		
' Buttons: When you click one of the buttons, the text that appears on the button is copied onto the users' clipboard to paste where
'		needed. Reduces the syntax errors found in BOMs.
'-----------------------------------------------------------------------------------------------------------------------------------------

Public Class PCADProcess

	Private Sub PCADProcess_Load() Handles MyBase.Load
		CenterToParent()
	End Sub

	Private Sub Close_Button_Click() Handles Close_Button.Click
		Close()
	End Sub

#Region "Process"
	Private Sub Handflow_Button_Click() Handles Handflow_Button.Click
		Clipboard.SetText(Handflow_Button.Text)
	End Sub

	Private Sub NotUsed_Button_Click() Handles NotUsed_Button.Click
		Clipboard.SetText(NotUsed_Button.Text)
	End Sub

	Private Sub PCBBoard_Button_Click() Handles PCBBoard_Button.Click
		Clipboard.SetText(PCBBoard_Button.Text)
	End Sub

	Private Sub PostAssembly_Button_Click() Handles PostAssembly_Button.Click
		Clipboard.SetText(PostAssembly_Button.Text)
	End Sub

	Private Sub SMT_Button_Click() Handles SMT_Button.Click
		Clipboard.SetText(SMT_Button.Text)
	End Sub

	Private Sub SMTBottom_Button_Click() Handles SMTBottom_Button.Click
		Clipboard.SetText(SMTBottom_Button.Text)
	End Sub

	Private Sub SMTHand_Button_Click() Handles SMTHand_Button.Click
		Clipboard.SetText(SMTHand_Button.Text)
	End Sub

	Private Sub BAS_Button_Click() Handles BAS_Button.Click
		Clipboard.SetText(BAS_Button.Text)
	End Sub
#End Region

#Region "Fields"
	Private Sub Vendor_Button_Click() Handles Vendor_Button.Click
		Clipboard.SetText(Vendor_Button.Text)
	End Sub

	Private Sub PartNumber_Button_Click() Handles PartNumber_Button.Click
		Clipboard.SetText(PartNumber_Button.Text)
	End Sub

	Private Sub StockNumber_Button_Click() Handles StockNumber_Button.Click
		Clipboard.SetText(StockNumber_Button.Text)
	End Sub

	Private Sub Process_Button_Click() Handles Process_Button.Click
		Clipboard.SetText(Process_Button.Text)
	End Sub

	Private Sub Option_Button_Click() Handles Option_Button.Click
		Clipboard.SetText(Option_Button.Text)
	End Sub

	Private Sub Swap_Button_Click() Handles Swap_Button.Click
		Clipboard.SetText(Swap_Button.Text)
	End Sub
#End Region

End Class