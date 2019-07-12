'-----------------------------------------------------------------------------------------------------------------------------------------
' Module: Variables.vb
'
' Description: Contains all of the common variables that are used in the entire project more than once as well as some debugging tools.
'
' Variable Discriptions:
'	_cn = This is the ODBC Connection that we use to access the Quickbooks database. It uses the active Quickbooks connection to preform
'		the ODBC queries. These queries are used to check if there has been an update to any of the items and gets item information and 
'		build assemblies for importing.
'
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
' Module: Login.vb
'
' Description: This is the opening window of the program. The user must have a valid user name and password to log in.
'
' Special Keys:
'   enter = enters the user name and password and trys to log the user in.
'-----------------------------------------------------------------------------------------------------------------------------------------
Imports System.Data.Odbc
Imports System.Data.SqlClient
Imports System.IO

'This module contains all of the common variables that are used in the entire project.
Public Module Variables
	'Enable logging
	Public Const LOGDATA As Boolean = True

	'Loggings save locations
	Public Const DATABASE As String = "Production"      'Production = BOMCompare	Devel = BOMDevel
	Public logLocation As String = "\\SERVER1\EngineeringReleased\Utilities\Compare Logs\"
	'Public logLocation As String = "C:\dev\Compare Logs\"

	Public myConn As SqlConnection
	Public sqlapi As New SQL_API()

	Public _cn As OdbcConnection = New OdbcConnection("DSN=Quickbooks Data;OLE DB Services=-2;")

	'Dilimiters
	Public Const PERIOD_DILIMITER As String = "."
	Public Const SPACE_DILIMITER As String = " "
	Public Const COMMA_DILIMITER As String = ","
	Public Const VLINE_DILIMITER As String = "|"

	'File Name parse index
	Public Const INDEX_BOARD As Integer = 0
	Public Const INDEX_REVISION1 As Integer = 1
	Public Const INDEX_REVISION2 As Integer = 2
	Public Const INDEX_OPTION As Integer = 3

	'Process'
	Public Const PROCESS_SMT As String = "SMT"
	Public Const PROCESS_SMTBOTTOM As String = "SMT BOTTOM"
	Public Const PROCESS_SMTHAND As String = "SMT HAND"
	Public Const PROCESS_HANDFLOW As String = "HAND FLOW"
	Public Const PROCESS_POSTASSEMBLY As String = "POST ASSEMBLY"
	Public Const PROCESS_PCBBOARD As String = "PCB BOARD"
	Public Const PROCESS_NOTUSED As String = "NOT USED"
	Public Const PROCESS_BAS As String = "BAS"

	'Table names
	Public Const TABLE_ALPHABOM As String = "ALPHA_BOM"
	Public Const TABLE_PCADBOM As String = "PCAD_BOM"
	Public Const TABLE_QBBOM As String = "QB_BOM"
	Public Const TABLE_QB_ITEMS As String = "QB_Items"
	Public Const TABLE_TEMP_PCADBOM As String = "Temp_PCAD_BOM"
	Public Const TABLE_ALPHA_ITEMS As String = "ALPHA_Items"
	Public Const TABLE_MAGAZINE_DATA As String = "Magazine_Data"
	Public Const TABLE_TEMP_PNP As String = "Temp_PNP"
	Public Const TABLE_ALPHA_PACKAGE As String = "ALPHA_Package"
	Public Const TABLE_UTILITIES As String = "Utilities"

	'Special Reference Designator
	Public Const REFERENCE_DESIGNATOR_OPTION As String = "ZD"
	Public Const REFERENCE_DESIGNATOR_FIDUCIAL As String = "ZF"
	Public Const REFERENCE_DESIGNATOR_SWAP As String = "ZX"

	'Special Prefixs
	Public Const PREFIX_PCB As String = "PCB"
	Public Const PREFIX_SMA As String = "SMA"
	Public Const PREFIX_BAS As String = "BAS"
	Public Const PREFIX_BIS As String = "BIS"
	Public Const PREFIX_DAS As String = "DAS"
	Public Const PREFIX_FGS As String = "FGS"

	'Special SQL Returns
	Public Const DOES_NOT_EXIST As String = "Does not exist"
	Public Const NOT_IN_DATABASE As String = "NOT IN DATABASE"

	'Extentions
	Public Const ALPHA_EXE As String = "*.gen"
	Public Const PCAD_EXE As String = "*.bom.csv"
	Public Const QB_EXE As String = "*.csv"
	Public Const QBITEMS_EXE = "*.csv"
	Public Const PNP_CSV_EXE As String = "*.pnp.csv"
	Public Const PCB_EXE As String = "*.pcb"
	Public Const SCH_EXE As String = "*.sch"
	Public Const SCH_PDF_EXE As String = "*.sch.pdf"
	Public Const ZIP_EXE As String = "*.zip"
	Public Const REV_TXT_EXE As String = "*.rev.txt"

	Public ODBC_ItemQuery As String = "SELECT " &
										"[ParentRefFullName] As """ & DB_HEADER_ITEM_PREFIX & """, " &
										"[Name] AS """ & DB_HEADER_ITEM_NUMBER & """," &
										"[Type] AS """ & DB_HEADER_TYPE & """, " &
										"[Description] AS """ & DB_HEADER_DESCRIPTION & """," &
										"[PrefVendorRefFullName] AS """ & DB_HEADER_VENDOR & """, " &
										"[ManufacturerPartNumber] AS """ & DB_HEADER_MPN & """," &
										"[QuantityOnHand] AS """ & DB_HEADER_QUANTITY & """," &
										"[PurchaseCost] AS """ & DB_HEADER_COST & """," &
										"[CustomFieldLeadTimeWeeks] AS """ & DB_HEADER_LEAD_TIME & """, " &
										"[CustomFieldMinimumOrderQTY] AS """ & DB_HEADER_MIN_ORDER_QTY & """," &
										"[CustomFieldReOrderQTY] AS """ & DB_HEADER_REORDER_QTY & """, " &
										"[CustomFieldManufacturer2] AS """ & DB_HEADER_VENDOR2 & """," &
										"[CustomFieldPartNumber2] AS """ & DB_HEADER_MPN2 & """, " &
										"[CustomFieldManufacturer3] AS """ & DB_HEADER_VENDOR3 & """," &
										"[CustomFieldPartNumber3] AS """ & DB_HEADER_MPN3 & """" &
										"From Item Where " &
										"[IsActive] = 1 And " &
										"[Type] <> 'ItemService' AND " &
										"[Type] <> 'ItemNonInventory' AND " &
										"[ParentRefFullName] <> ''"

	Public ODBC_AssemblyQuery As String = "SELECT " &
										"[ParentRefFullName] AS """ & DB_HEADER_NAME_PREFIX & """," &
										"[Name] As """ & DB_HEADER_NAME & """, " &
										"{fn SUBSTRING([ItemInventoryAssemblyLnItemInventoryRefFullName], 1, {fn LOCATE(':',[ItemInventoryAssemblyLnItemInventoryRefFullName])} - 1 )} AS """ & DB_HEADER_ITEM_PREFIX & """, " &
										"{fn SUBSTRING([ItemInventoryAssemblyLnItemInventoryRefFullName], {fn LOCATE(':',[ItemInventoryAssemblyLnItemInventoryRefFullName])} + 1, 1000)} AS """ & DB_HEADER_ITEM_NUMBER & """, " &
										"[ItemInventoryAssemblyLNQuantity] AS """ & DB_HEADER_QUANTITY & """ " &
										"FROM itemInventoryAssemblyLine " &
										"WHERE " &
										"[IsActive] = 1"

	Public ODBC_LastModifiedQuery As String = "SELECT TimeModified FROM Item UNOPTIMIZED WHERE " &
										"[IsActive] = 1 And " &
										"[Type] <> 'ItemService' AND " &
										"[Type] <> 'ItemNonInventory' AND " &
										"[ParentRefFullName] <> '' ORDER BY TimeModified DESC"

#Region "Table Headers"
	Public Const HEADER_REMAINDER As String = "Remainder"
	Public Const HEADER_TOTAL As String = "Total"
	Public Const HEADER_QTY_PCAD As String = "PCAD quantity"
	Public Const HEADER_QTY_QB As String = "QB quantity"
	Public Const HEADER_QTY_ORIG As String = "Orig Qty"
	Public Const HEADER_QTY_NEEDED As String = "Qty Needed"
	Public Const HEADER_QTY_AVAIL As String = "Avail Qty"
#End Region

#Region "Database Headers"
	Public Const DB_HEADER_NAME As String = "Name"                                          'Utilities |          |				  |          | QB_BOM |		     | Magazine_Data |			   |		   |
	Public Const DB_HEADER_VALUE As String = "Value"                                        'Utilities |		  |				  |			 |		  |		     |			     |			   |		   |
	Public Const DB_HEADER_REF_DES As String = "Ref Des"                                    '		   | Temp_PNP | Temp_PCAD_BOM |			 |		  | PCAD_BOM |			     |			   | ALPHA_BOM |
	Public Const DB_HEADER_ITEM_PREFIX As String = "Item Prefix"                            '		   | Temp_PNP | Temp_PCAD_BOM | QB_Items | QB_BOM | PCAD_BOM |			     |			   |		   |
	Public Const DB_HEADER_ITEM_NUMBER As String = "Item Number"                            '		   | Temp_PNP | Temp_PCAD_BOM | QB_Items | QB_BOM | PCAD_BOM | Magazine_Data | ALPHA_Items | ALPHA_BOM |
	Public Const DB_HEADER_POS_X As String = "Pos X"                                        '		   | Temp_PNP |				  |			 |		  |		     |			     |			   | ALPHA_BOM |
	Public Const DB_HEADER_POS_Y As String = "Pos Y"                                        '		   | Temp_PNP |				  |			 |		  |		     |			     |			   | ALPHA_BOM |
	Public Const DB_HEADER_ROTATION As String = "Rotation"                                  '		   | Temp_PNP |				  |			 |		  |		     |			     |			   |		   |
	Public Const DB_HEADER_PROCESS As String = "Process"                                    '		   | Temp_PNP | Temp_PCAD_BOM |			 | QB_BOM | PCAD_BOM |			     |			   | ALPHA_BOM |
	Public Const DB_HEADER_BOARD_NAME As String = "Board Name"                              '		   |		  | Temp_PCAD_BOM |			 |		  | PCAD_BOM |			     |			   | ALPHA_BOM |
	Public Const DB_HEADER_DESCRIPTION As String = "Description"                            '		   |		  | Temp_PCAD_BOM | QB_Items |		  | PCAD_BOM |			     |			   |		   |
	Public Const DB_HEADER_VENDOR As String = "Vendor"                                      '		   |		  | Temp_PCAD_BOM | QB_Items |		  | PCAD_BOM |			     |			   |		   |
	Public Const DB_HEADER_MPN As String = "MPN"                                            '		   |		  | Temp_PCAD_BOM | QB_Items |		  | PCAD_BOM |			     |			   |		   |
	Public Const DB_HEADER_OPTION As String = "Option"                                      '		   |		  | Temp_PCAD_BOM |			 |		  |		     |			     |			   |		   |
	Public Const DB_HEADER_SWAP As String = "Swap"                                          '		   |		  | Temp_PCAD_BOM |			 |		  |		     |			     |			   |		   |
	Public Const DB_HEADER_ERRORS As String = "Errors"                                      '		   |		  | Temp_PCAD_BOM |			 |		  |		     |			     |			   |		   |
	Public Const DB_HEADER_TYPE As String = "Type"                                          '		   |		  |				  | QB_Items | QB_BOM |		     |			     |			   |		   |
	Public Const DB_HEADER_QUANTITY As String = "Quantity"                                  '		   |		  |				  | QB_Items | QB_BOM |		     | Magazine_Data |			   |		   |
	Public Const DB_HEADER_COST As String = "Cost"                                          '		   |		  |				  | QB_Items | QB_BOM |		     |			     |			   |		   |
	Public Const DB_HEADER_VENDOR2 As String = "Vendor 2"                                   '		   |		  |				  | QB_Items |		  |		     |			     |			   |		   |
	Public Const DB_HEADER_MPN2 As String = "MPN 2"                                         '		   |		  |				  | QB_Items |		  |		     |			     |			   |		   |
	Public Const DB_HEADER_VENDOR3 As String = "Vendor 3"                                   '		   |		  |				  | QB_Items |		  |		     |			     |			   |		   |
	Public Const DB_HEADER_MPN3 As String = "MPN 3"                                         '		   |		  |				  | QB_Items |		  |		     |			     |			   |		   |
	Public Const DB_HEADER_LEAD_TIME As String = "Lead Time"                                '		   |		  |				  | QB_Items |		  |		     |			     |			   |		   |
	Public Const DB_HEADER_MIN_ORDER_QTY As String = "Min Order Qty"                        '		   |		  |				  | QB_Items |		  |		     |			     |			   |		   |
	Public Const DB_HEADER_REORDER_QTY As String = "ReOrder Qty"                            '		   |		  |				  | QB_Items |		  |		     |			     |			   |		   |
	Public Const DB_HEADER_NAME_PREFIX As String = "Name Prefix"                            '		   |		  |				  |			 | QB_BOM |		     | Magazine_Data |			   |		   |
	Public Const DB_HEADER_SERIAL_NUMBER As String = "Serial Number"                        '		   |		  |				  |			 |		  |		     | Magazine_Data |			   |		   |
	Public Const DB_HEADER_SLOT_NUMBER As String = "Slot Number"                            '		   |		  |				  |			 |		  |		     | Magazine_Data |			   |		   |
	Public Const DB_HEADER_MACHINE_NUMBER As String = "Machine Number"                      '		   |		  |				  |			 |		  |		     | Magazine_Data |			   |		   |
	Public Const DB_HEADER_FEEDER_NUMBER As String = "Feeder Number"                        '		   |		  |				  |			 |		  |		     | Magazine_Data |			   |		   |
	Public Const DB_HEADER_ANGLE As String = "Angle"                                        '		   |		  |				  |			 |		  |		     | Magazine_Data |			   | ALPHA_BOM |
	Public Const DB_HEADER_FEEDER_TYPE As String = "Feeder Type"                            '		   |		  |				  |			 |		  |		     |			     | ALPHA_Items |		   |
	Public Const DB_HEADER_MAGAZINE_TYPE As String = "Magazine Type"                        '		   |		  |				  |			 |		  |		     |			     | ALPHA_Items |		   |
	Public Const DB_HEADER_DEFAULT_TAPE_ANGLE As String = "Default Tape Angle"              '		   |		  |				  |			 |		  |		     |			     | ALPHA_Items |		   |
	Public Const DB_HEADER_PACKAGE As String = "Package"                                    '		   |		  |				  |			 |		  |		     |			     | ALPHA_Items |		   |
	Public Const DB_HEADER_DEFAULT_STEP_LENGTH As String = "Default Step Length"            '		   |		  |				  |			 |		  |		     |			     | ALPHA_Items |		   |
	Public Const DB_HEADER_DEFAULT_STEP_LENGTH_TRIM As String = "Default Step Length Trim"  '		   |		  |				  |			 |		  |		     |			     | ALPHA_Items |		   |
	Public Const DB_HEADER_DEFAULT_STEPS As String = "Default Steps"                        '		   |		  |				  |			 |		  |		     |			     | ALPHA_Items |		   |
	Public Const DB_HEADER_FINEPITCH As String = "Finepitch"                                '		   |		  |				  |			 |		  |		     |			     | ALPHA_Items |		   |
	Public Const DB_HEADER_GROUP As String = "Group"                                        '		   |		  |				  |			 |		  |		     |			     |			   | ALPHA_BOM |
	Public Const DB_HEADER_MOUNT_SKIP As String = "Mount-Skip"                              '		   |		  |				  |			 |		  |		     |			     |			   | ALPHA_BOM |
	Public Const DB_HEADER_DESPENSE_SKIP As String = "Despense-Skip"                        '		   |		  |				  |			 |		  |		     |			     |			   | ALPHA_BOM |

	'ALPHA_Package
	Public Const DB_HEADER_PACKAGE_NAME As String = "Package Name"
	Public Const DB_HEADER_BODY_LENGTH As String = "Body Length"
	Public Const DB_HEADER_BODY_WIDTH As String = "Body Width"
	Public Const DB_HEADER_OVERALL_LENGTH As String = "Overall Length"
	Public Const DB_HEADER_OVERALL_WIDTH As String = "Overall Width"
	Public Const DB_HEADER_NORMAL_HEIGHT As String = "Normal Height"
	Public Const DB_HEADER_MAX_HEIGHT As String = "Max Height"
	Public Const DB_HEADER_MIN_HEIGHT As String = "Min Height"
	Public Const DB_HEADER_ANGLE_1 As String = "Angle 1"
	Public Const DB_HEADER_LEVEL_1 As String = "Level 1"
	Public Const DB_HEADER_POSITION_1 As String = "Position 1"
	Public Const DB_HEADER_FORCE_1 As String = "Force 1"
	Public Const DB_HEADER_NORMAL_SIZE_1 As String = "Normal Size 1"
	Public Const DB_HEADER_MAX_SIZE_1 As String = "Max Size 1"
	Public Const DB_HEADER_MIN_SIZE_1 As String = "Min Size 1"
	Public Const DB_HEADER_VERIFY_MECHANICAL_1 As String = "Verify Mechanical 1"
	Public Const DB_HEADER_ANGLE_2 As String = "Angle 2"
	Public Const DB_HEADER_LEVEL_2 As String = "Level 2"
	Public Const DB_HEADER_POSITION_2 As String = "Position 2"
	Public Const DB_HEADER_FORCE_2 As String = "Force 2"
	Public Const DB_HEADER_NORMAL_SIZE_2 As String = "Normal Size 2"
	Public Const DB_HEADER_MAX_SIZE_2 As String = "Max Size 2"
	Public Const DB_HEADER_MIN_SIZE_2 As String = "Min Size 2"
	Public Const DB_HEADER_VERIFY_MECHANICAL_2 As String = "Verify Mechanical 2"
#End Region

	'This function is designed to log to a CSV file when we have experienced an issue that has not been handled yet.
	'This is to help find, debug, and fix the problem.
	Public Sub Log(ByRef Operation As String, ByRef Details As String)
		'Check to see if we are going to log the data or not.
		If LOGDATA = False And Not Operation Is Nothing Then
			Return
		End If

		Try
			Dim path As String = String.Format("{0}Log_{1}.csv", logLocation, DateTime.Today.ToString("yyyy-MM-dd"))
			If Not (File.Exists(path)) Then
				File.AppendAllText(path, String.Format("Time,User,Operation,Version,Details{0}", vbNewLine))
			End If

			If Operation Is Nothing Then
				File.AppendAllText(path, String.Format(",,,,{0}{1}", Details, vbNewLine))
			Else
				File.AppendAllText(path, String.Format("{0},{1},{2},{3},{4}{5}", DateTime.Now, sqlapi._Username, Operation, Application.ProductVersion, Details, vbNewLine))
			End If
		Catch ex As Exception
			'MsgBox("LOG ERROR: " & ex.Message)
		End Try
	End Sub

	Public Sub UnhandledExceptionMessage(ByRef ex As Exception)
		Dim temp As String = ex.StackTrace.Replace(vbCr, "").Replace(vbLf, "").Replace("at ", "|").Replace(",", ";")
		Dim split() As String = temp.Split("|")
		Log(ex.TargetSite.DeclaringType.Name & ":" & ex.TargetSite.Name, ex.Message)
		For index As Integer = 1 To split.Length - 1
			Log(Nothing, split(index).Trim)
		Next
		MessageBox.Show("Congrats " & sqlapi._Username & "! You were able to crash the program!" & vbNewLine &
						"Hopefully I got enough data logged to fix it but just to make sure," & vbNewLine &
						"try to remember how you managed it and try to crash it again." & vbNewLine &
						"See " & logLocation & " for more details." & vbNewLine &
						"Recomend closing current window or program to prevent unsteady state." & vbNewLine &
						ex.Message, "Error")
	End Sub

	Public Sub RefreshMags()
		'Refresh the magazine data if there is a newer file found.
		Dim message As String = ""
		If sqlapi.CheckDirtyBit(message) = True Then
			MsgBox(message)
			Return
		End If

		'We need to set the dirty bit in the database because we could be preforming an import.
		sqlapi.SetDirtyBit(1)

		Dim newInfo As FileInfo
		Dim oldinfo As FileInfo

		'Check to see if we have a new file
		If File.Exists("\\Server1\Shares\Production\AlphaBackup\mag.mag") Then
			newInfo = New FileInfo("\\Server1\Shares\Production\AlphaBackup\mag.mag")
		Else
			'There is no file found in AlphaBackup so we cannot refresh the mag info.
			MsgBox("Could not find file: \\Server1\Shares\Production\AlphaBackup\mag.mag")
			Return
		End If

		'Check to make sure that we have a mag file in our released location
		If File.Exists(My.Settings.ReleaseLocation & "\ALPHAITEMS\mag.mag") Then
			oldinfo = New FileInfo(My.Settings.ReleaseLocation & "\ALPHAITEMS\mag.mag")

			If newInfo.LastWriteTime > oldinfo.LastWriteTime Then
				'We need to replace the file in SERVER1 with the new file in AlphaBackup

				File.Copy("\\Server1\Shares\Production\AlphaBackup\mag.mag", My.Settings.ReleaseLocation & "\ALPHAITEMS\mag.mag", True)
			Else
				'The file in SERVER1 is already the most up to date file.
				MsgBox("The Mag info is already up to date.")
				Return
			End If
		Else
			'The SERVER1 file does not exist so we just need to copy the file over.
			File.Copy("\\Server1\Shares\Production\AlphaBackup\mag.mag", My.Settings.ReleaseLocation & "\ALPHAITEMS\mag.mag", True)
		End If

		'Now we need to re-import the magazine information
		Dim myCmd As New SqlCommand("", myConn)
		myCmd.CommandText = "DELETE FROM " & TABLE_MAGAZINE_DATA
		myCmd.ExecuteNonQuery()

		Dim form As New ImportData(Nothing)
		Dim result As String = ""
		If form.ImportALPHAmagazine(My.Settings.ReleaseLocation & "\ALPHAITEMS", "mag.mag", result) = False Then
			MsgBox("Something happened while trying to update magizine information: " & result)
		End If

		File.Delete("\\Server1\Shares\Production\AlphaBackup\mag.mag")

		MsgBox("Magizine information has been updated. Generate your report again for a more accurate result.")
	End Sub

	Public Function ChangeCheck(ByRef report As Boolean) As Boolean
		Dim hasChanged = False
		Dim importTime As Date
		Dim databasetime As Date

		If My.Settings.RemindMe = False Then
			Return True
		End If

		If My.Settings.DatabasePath.Length <> 0 Then
			If File.Exists(My.Settings.DatabasePath) Then
				Dim databaseFileInfo = New FileInfo(My.Settings.DatabasePath)
				databasetime = databaseFileInfo.LastWriteTime

				'Get the last time the database was updated.
				Dim myCmd = New SqlCommand("SELECT [" & DB_HEADER_VALUE & "] FROM " & TABLE_UTILITIES & " WHERE [" & DB_HEADER_NAME & "] = 'LastUpdate'", myConn)
				importTime = Date.Parse(myCmd.ExecuteScalar)

				If databasetime > importTime Then
					hasChanged = True
				End If
			Else
				Dim alert As New MessageBoxQB("Database file at '" & My.Settings.DatabasePath & "' does not exist. If you continue you might be using old data." & vbNewLine &
											  vbNewLine &
											  "DETAILS: Could not check for update." & vbNewLine &
											  "Make sure the file exists at the location being used or fix the path in the settings.",
											  "Continue", "Cancel", Nothing, True)

				If alert.ShowDialog() = DialogResult.No Then
					Return False
				End If

				Return True
			End If
		Else
			Dim alert As New MessageBoxQB("Database Path has not been defined yet. If you continue you might be using old data." & vbNewLine &
										  vbNewLine &
										  "DETAILS: Could not check for update." & vbNewLine &
										  "Make sure you set up the path to the database file in the settings.",
										  "Continue", "Cancel", Nothing, True)

			If alert.ShowDialog() = DialogResult.No Then
				Return False
			End If

			Return True
		End If

		If hasChanged = True Then
			'Open a connection and see if our items list has changed or not
			Try
				_cn.Open()
			Catch ex As OdbcException
				Dim alert As New MessageBoxQB("Could not connect to QB database to check for changes. If you continue you might be using old data. " & vbNewLine &
											  vbNewLine &
											  "DETAILS: " & ex.Errors.Item(0).Message & vbNewLine &
											  vbNewLine &
											  "Make sure you have QB running on your computer and you have granted access for this program. If both of these are already done then read the DETAILS to see what error is being thrown.", "Continue", "Cancel", Nothing, True)

				If alert.ShowDialog() = DialogResult.No Then
					Return False
				End If

				Return True
			End Try

			Dim cmd As New OdbcCommand(ODBC_LastModifiedQuery, _cn)
			databasetime = Date.Parse(cmd.ExecuteScalar)
			_cn.Close()

			'Check to see if we have a newer Item modified time.
			If databasetime > importTime Then
				Dim alert As New MessageBoxQB("QB Database items have been changed. If you continue you could be using old data." & vbNewLine &
											  "Would you like to Continue, Cancel, or Import?", "Continue", "Cancel", "Import", True)

				If alert.ShowDialog() = DialogResult.No Then
					Return False
				ElseIf alert.DialogResult = DialogResult.Ignore Then
					MenuMain.OpenForm(New ImportData(MenuMain))
					Return False
				End If

				Return True
			End If
		End If

		Return True
	End Function

	Public Sub HasDatabaseChanged()
		If My.Settings.RemindMe = False Then
			Return
		End If

		If My.Settings.DatabasePath.Length = 0 Then
			Dim alert As New MessageBoxQB("Database Path has not been defined yet. Recomend going to the settings and updating this information.", "OK", Nothing, Nothing, True)

			If alert.ShowDialog() = DialogResult.Yes Then
				'Do nothing here because this should allow the code here to pause.
			End If
			Return
		End If

		If File.Exists(My.Settings.DatabasePath) Then
			'Get the database file last write time.
			Dim databaseFileInfo = New FileInfo(My.Settings.DatabasePath)
			Dim databaseTime As Date = databaseFileInfo.LastWriteTime

			'Get our last import time.
			'See if we already have a lastUpdate in our database records.
			Dim myCmd As New SqlCommand("IF EXISTS(SELECT [" & DB_HEADER_NAME & "] FROM " & TABLE_UTILITIES & " WHERE [" & DB_HEADER_NAME & "] = 'LastUpdate') SELECT 1 ELSE SELECT 0", myConn)

			If myCmd.ExecuteScalar = 0 Then
				myCmd = New SqlCommand("INSERT INTO " & TABLE_UTILITIES & " ([" & DB_HEADER_NAME & "], [" & DB_HEADER_VALUE & "]) VALUES('LastUpdate','')", myConn)
				myCmd.ExecuteNonQuery()
			End If

			myCmd = New SqlCommand("SELECT [" & DB_HEADER_VALUE & "] FROM " & TABLE_UTILITIES & " WHERE [" & DB_HEADER_NAME & "] = 'LastUpdate'", myConn)

			Dim importTime As Date = Date.Parse(myCmd.ExecuteScalar)

			'Check to see if we have a newer database time.
			If databaseTime > importTime Then
				'Open a connection and see if our items list has changed or not
				Try
					_cn.Open()
				Catch ex As OdbcException
					Dim alert As New MessageBoxQB("Could not connect to QB database to check for changes. If you continue without importing you might be using old data. " & vbNewLine &
												  vbNewLine &
												  "DETAILS: " & ex.Errors.Item(0).Message & vbNewLine &
												  vbNewLine &
												  "Make sure you have QB running on your computer and you have granted access for this program. If both of these are already done then read the DETAILS to see what error is being thrown.", "OK", Nothing, Nothing, True)

					If alert.ShowDialog() = DialogResult.Yes Then
						'Do nothing here because this should allow the code here to pause.
					End If
					Return
				End Try

				Dim cmd As New OdbcCommand(ODBC_LastModifiedQuery, _cn)
				databaseTime = Date.Parse(cmd.ExecuteScalar)
				_cn.Close()

				'Check to see if we have a newer Item modified time.
				If databaseTime > importTime Then
					'We need to call a special window that askes the user if they want to update the database.
					Dim alert As New MessageBoxQB("The QB Database has been updated recently. Would you like to run an import right now?", "Yes", "No", Nothing, True)

					If alert.ShowDialog() = DialogResult.Yes Then
						MenuMain.OpenForm(New ImportData(MenuMain))
					End If
				End If
			End If
		Else
			Dim alert As New MessageBoxQB("Database file at '" & My.Settings.DatabasePath & "' does not exist. If you continue without fixing the path you might be using old data." & vbNewLine &
											  vbNewLine &
											  "DETAILS: Could not check for update." & vbNewLine &
											  "Make sure the file exists at the location being used or fix the path in the settings.",
											  "OK", Nothing, Nothing, True)

			If alert.ShowDialog() = DialogResult.Yes Then
				'Do nothing here because this should allow the code here to pause.
			End If
			Return
		End If
	End Sub

End Module

Public Class Login

	Private Sub Login_Load() Handles MyBase.Load
		L_Version.Text = "V:" & Application.ProductVersion
		Database_Label.Text = DATABASE
		KeyPreview = True

		''Reset our remind flag back to true for this session until the user turns it off.
		'My.Settings.RemindMe = True
		'My.Settings.Save()
	End Sub

	Private Sub MyBase_KeyDown(ByVal sender As Object, ByVal e As KeyEventArgs) Handles MyBase.KeyDown
		'Call the Login Function when we press the Enter key.
		If e.KeyCode.Equals(Keys.Enter) Then
			Call B_Login_Click()
		End If
	End Sub

	Private Sub B_Login_Click() Handles B_Login.Click
		Dim myCmd As New SqlCommand
		sqlapi._Username = TB_User.Text
		sqlapi._Password = TB_Password.Text

		'Try to open the database. If we cannot, then we have the wrong username or password.
		If sqlapi.OpenDatabase(myConn, myCmd) = False Then
			TB_Password.SelectAll()
			Return
		End If

		Dim DoMenuMain As New MenuMain
		DoMenuMain.Show()
		Close()
	End Sub

	Private Sub B_Exit_Click() Handles B_Exit.Click
		Application.Exit()
	End Sub

End Class