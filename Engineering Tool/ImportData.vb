'-----------------------------------------------------------------------------------------------------------------------------------------
' Module: ImportData.vb
'
' Description: Imports all of the data from our Release directory into our database. There are a lot of syntax, naming, fields, and
'		duplicate checking. Any errors are put to the screen for the user to see. Light indicators on the left side indicate and navigate
'		where the errors are. 
'
' Checkboxes:
'	Import ALPHA Items = Include the importation of ALPHA Items, Magazine, and Package data during the import. Optional because this is
'		not used very often and will take the import longer to complete.
'
' Light status:
'	Green = This section had no issues and everything was imported to the database.
'	Yellow = This section had a few minor issues but everything was still imported to the database.
'	Red = This section had major issues. These items will not have been imported to the database.
'-----------------------------------------------------------------------------------------------------------------------------------------
Imports System.Data.SqlClient
Imports System.IO
Imports Microsoft.VisualBasic.FileIO
Imports System.Data.Odbc

Public Class ImportData

	'Exceptions for Manufacturer/Part Number
	Const TYPE_INVENTORY_ASSEMBLY As String = "Inventory Assembly"
	Const TYPE_INVENTORY_PART As String = "Inventory Part"

	'ALPHA BOM Parses
	Const ALHPA_CODE_INDEX As Integer = 0
	Const ALPHA_NAME_INDEX As Integer = 1
	Const ALHPA_REFERENCE_DESIGNATOR_INDEX As Integer = 1
	Const ALHPA_X_POSITION_INDEX As Integer = 1
	Const ALHPA_Y_POSITION_INDEX As Integer = 2
	Const ALHPA_ANGLE_INDEX As Integer = 3
	Const ALHPA_GROUP_INDEX As Integer = 4
	Const ALHPA_MOUNT_SKIP_INDEX As Integer = 5
	Const ALHPA_DESPENSE_SKIP_INDEX As Integer = 6
	Const ALHPA_COMPONENT_INDEX As Integer = 7

	Private _MenuMain As MenuMain
	Dim needsPNP As Boolean = False

	Dim skipQB As Boolean = False

	'Error types that we have encountered. 
	Dim ALPHAmajor As Boolean = False
	Dim ALPHAminor As Boolean = False

	Dim ALPHAitemsMajor As Boolean = False
	Dim ALPHAitemsMinor As Boolean = False

	Dim PCADmajor As Boolean = False
	Dim PCADminor As Boolean = False

	Dim QBmajor As Boolean = False
	Dim QBminor As Boolean = False

	Dim QBitemsMajor As Boolean = False
	Dim QBitemsMinor As Boolean = False

	Dim myCmd As New SqlCommand("", myConn)

	Public Sub New(ByRef mainform As MenuMain)
		InitializeComponent()

		'We need this so we can update the text on the main page if we do an update.
		_MenuMain = mainform
	End Sub

	Private Sub Import_Button_Click() Handles Import_Button.Click
		Cursor = Cursors.WaitCursor
		If LOGDATA = True Then
			Try
				Import()
			Catch ex As Exception
				UnhandledExceptionMessage(ex)
			End Try
		Else
			Import()
		End If
		Cursor = Cursors.Default
	End Sub

	Private Sub Import()

		Dim answer As DialogResult

		'This check allows the user to reset the flag that tells us if an import is going on. This way if the program crashes, we have a way to recover.
		Dim message As String = ""
		If sqlapi.CheckDirtyBit(message) = True Then
			answer = MessageBox.Show(message & vbNewLine & "Would you like to reset the flag??", "Reset?", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)
			If answer = Windows.Forms.DialogResult.Yes Then
				sqlapi.SetDirtyBit(0)
			End If
			Return
		End If

		'This is a safegaurd agaisnt any kind of accidental imports. The user needs to verify that they want to continue.
		
		answer = MessageBox.Show("Are you sure you want to start importing data?", "Continue?", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)
		If answer = Windows.Forms.DialogResult.No Then
			Return
		End If

		skipQB = False

		'Dim stopwatch As Stopwatch = Stopwatch.StartNew()

		'Clear our textbox to start with a clean slate.
		RTB_Results.Clear()

		Try
			UpdateRTB("Opening QB Connection . . ." & vbNewLine)

			_cn.Open()

			UpdateRTB("Connected!" & vbNewLine & vbNewLine)
		Catch ex As Exception
			answer = MessageBox.Show("Could not connect to QB. Continue anyways?", "Continue?", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)
			If answer = Windows.Forms.DialogResult.No Then
				UpdateRTB("Could not connect and canceled." & vbNewLine)
				Return
			End If

			UpdateRTB("Could not connect." & vbNewLine & vbNewLine)
			skipQB = True
		End Try
		'stopwatch.Stop()

		'Check to see if the directory exists.
		Dim missingDirectory As String = ""
		Dim isMissingDirectory As Boolean = False

		If Directory.Exists(My.Settings.ReleaseLocation & "\ALPHABOM") = False Then
			missingDirectory = missingDirectory & My.Settings.ReleaseLocation & "\ALPHABOM" & vbNewLine
			isMissingDirectory = True
		End If

		If Directory.Exists(My.Settings.ReleaseLocation & "\ALPHAITEMS") = False Then
			missingDirectory = missingDirectory & My.Settings.ReleaseLocation & "\ALPHAITEMS" & vbNewLine
			isMissingDirectory = True
		End If

		If Directory.Exists(My.Settings.ReleaseLocation & "\PCAD") = False Then
			missingDirectory = missingDirectory & My.Settings.ReleaseLocation & "\PCAD" & vbNewLine
			isMissingDirectory = True
		End If

		If isMissingDirectory = True Then
			MsgBox("Release location " & My.Settings.ReleaseLocation & " does not contain the following directories required for import:" & vbNewLine & missingDirectory)
			Return
		End If

		Dim cmd As New OdbcCommand(ODBC_LastModifiedQuery, _cn)

		Dim myCmd = New SqlCommand("SELECT [" & DB_HEADER_VALUE & "] FROM " & TABLE_UTILITIES & " WHERE [" & DB_HEADER_NAME & "] = 'LastUpdate'", myConn)

		Try
			'Set our dirty bit to flag others that an import is happening.
			sqlapi.SetDirtyBit(1)

			'Reset all of our error flags.
			ALPHAmajor = False
			ALPHAminor = False

			ALPHAitemsMajor = False
			ALPHAitemsMinor = False

			PCADmajor = False
			PCADminor = False

			QBmajor = False
			QBminor = False

			QBitemsMajor = False
			QBitemsMinor = False

			'UpdateRTB("Connection = " & stopwatch.Elapsed.ToString & vbNewLine)

			Dim result As String = ""
			If sqlapi.ClearDatabase(myCmd, result, skipQB, CkB_ALPHAitems.Checked) = False Then
				UpdateRTB(result & vbNewLine)
				Return
			End If

			'Change all of our 'Lights' to a basic netural color to show that we are in the import stage.
			TB_ALPHAIndicatorLight.BackColor = Color.White
			TB_ALPHAIndicatorLight.Refresh()

			TB_ALPHAitemsIndicatorLight.BackColor = Color.White
			TB_ALPHAitemsIndicatorLight.Refresh()

			TB_PCADIndicatorLight.BackColor = Color.White
			TB_PCADIndicatorLight.Refresh()

			TB_QBIndicatorLight.BackColor = Color.White
			TB_QBIndicatorLight.Refresh()

			TB_QBitemsIndicatorLight.BackColor = Color.White
			TB_QBitemsIndicatorLight.Refresh()

			ImportFilesToDatabase()

			_cn.Close()
			'Check to see if we have any errors. If we do then change the related light to the level of error that we have.
			If ALPHAmajor Then
				TB_ALPHAIndicatorLight.BackColor = Color.Red
			ElseIf ALPHAminor Then
				TB_ALPHAIndicatorLight.BackColor = Color.Yellow
			Else
				TB_ALPHAIndicatorLight.BackColor = Color.LightGreen
			End If

			If ALPHAitemsMajor Then
				TB_ALPHAitemsIndicatorLight.BackColor = Color.Red
			ElseIf ALPHAitemsMinor Then
				TB_ALPHAitemsIndicatorLight.BackColor = Color.Yellow
			Else
				TB_ALPHAitemsIndicatorLight.BackColor = Color.LightGreen
			End If

			If PCADmajor Then
				TB_PCADIndicatorLight.BackColor = Color.Red
			ElseIf PCADminor Then
				TB_PCADIndicatorLight.BackColor = Color.Yellow
			Else
				TB_PCADIndicatorLight.BackColor = Color.LightGreen
			End If

			If QBmajor Then
				TB_QBIndicatorLight.BackColor = Color.Red
			ElseIf QBminor Then
				TB_QBIndicatorLight.BackColor = Color.Yellow
			Else
				TB_QBIndicatorLight.BackColor = Color.LightGreen
			End If

			If QBitemsMajor Then
				TB_QBitemsIndicatorLight.BackColor = Color.Red
			ElseIf QBitemsMinor Then
				TB_QBitemsIndicatorLight.BackColor = Color.Yellow
			Else
				TB_QBitemsIndicatorLight.BackColor = Color.LightGreen
			End If

			SaveOutput_Button.Enabled = True

			'Update our lastUpdate time in the database.
			myCmd = New SqlCommand("UPDATE " & TABLE_UTILITIES & " SET [" & DB_HEADER_VALUE & "] = GETDATE() WHERE [" & DB_HEADER_NAME & "] = 'LastUpdate'", myConn)
			myCmd.ExecuteNonQuery()

			myCmd = New SqlCommand("SELECT [" & DB_HEADER_VALUE & "] FROM " & TABLE_UTILITIES & " WHERE [" & DB_HEADER_NAME & "] = 'LastUpdate'", myConn)
			_MenuMain.L_LastImport.Text = myCmd.ExecuteScalar
			_MenuMain.L_LastImport.Refresh()

			'If we also updated our Alpha files then update the lastALPHAUpdate time in the database as well.
			If CkB_ALPHAitems.Checked = True Then
				myCmd = New SqlCommand("UPDATE " & TABLE_UTILITIES & " SET [" & DB_HEADER_VALUE & "] = GETDATE() WHERE [" & DB_HEADER_NAME & "] = 'LastALPHAUpdate'", myConn)
				myCmd.ExecuteNonQuery()

				myCmd = New SqlCommand("SELECT [" & DB_HEADER_VALUE & "] FROM " & TABLE_UTILITIES & " WHERE [" & DB_HEADER_NAME & "] = 'LastALPHAUpdate'", myConn)
				_MenuMain.L_LastALPHAImport.Text = myCmd.ExecuteScalar
				_MenuMain.L_LastALPHAImport.Refresh()
			End If

			'Set the dirty bit to show that we had a successful import.
			sqlapi.SetDirtyBit(0)
			Log("Import", "Success")
		Catch ex As Exception
			UpdateRTB("ERROR ERROR ERROR ERROR: " & ex.Message & vbNewLine)
			Log("Import", "ERROR: " & ex.Message)
		End Try
	End Sub

	Private Sub SaveOutput_Button_Click() Handles SaveOutput_Button.Click
		Dim report As New GenerateReport()
		report.GenerateImportReport(RTB_Results)
	End Sub

	Private Sub Close_Button_Click() Handles Close_Button.Click
		Close()
	End Sub

	Private Sub ImportFilesToDatabase()
		Try
			'Go through each of the directories found in the passed in location. The following folder structure is expected as follows:
			'ROOT DIRECTORY [Boards]					~ This folder is set up inside the settings
			'	ALPHABOM								~ This folder needs to be exact
			'		<boardname>						There can be multiple board folders
			'			Rev#.#						There can be multiple revision folders
			'				files					Name format: <boardname>.Rev#.#.<option>.gen
			'	PCAD									~ This folder needs to be exact
			'		<boardname>						There can be multiple board folders
			'			Rev#.#						There can be multiple revision folders
			'				Released				Folder that contains all of the files for release
			'				files					Name format: <boardname>.Rev#.#.<option>.[ext]			should contain [.bom.csv][.pcb][.pnp.csv][.sch][.sch.pdf]

			For Each dir As String In Directory.GetDirectories(My.Settings.ReleaseLocation)
				Dim dirInfo As New DirectoryInfo(dir)

				'Find the folder names and go through them one by one.
				Select Case dirInfo.Name
					Case "ALPHABOM"
						UpdateRTB("Starting to import SMT BOM" & vbNewLine)

						GoThroughALPHABoards(dir)

						UpdateRTB("Finished importing SMT BOM" & vbNewLine & vbNewLine)
					Case "PCAD"
						UpdateRTB("Starting to import PCAD BOM" & vbNewLine)

						GoThroughImportPCADBoards(dir)

						UpdateRTB("Finished importing PCAD BOM" & vbNewLine & vbNewLine)
				End Select
			Next

			'Check to see if we are going to import the alpha files
			If CkB_ALPHAitems.Checked = True Then
				UpdateRTB("Starting to import SMT Items" & vbNewLine)

				GoThroughALPHAItems()

				UpdateRTB("Finished importing SMT Items" & vbNewLine & vbNewLine)
			Else
				UpdateRTB("Skipping The import of SMT Items" & vbNewLine & vbNewLine)
			End If

			If skipQB = False Then
				UpdateRTB("Starting to import QB Items" & vbNewLine)

				ImportQBItems()

				UpdateRTB("Finished importing QB Items" & vbNewLine & vbNewLine)

				'Import the QB BOMs
				UpdateRTB("Starting to import QB BOMs" & vbNewLine)

				ImportQBBOMS()

				UpdateRTB("Finished importing QB BOMs" & vbNewLine & vbNewLine)
			Else
				UpdateRTB("Skipping the import of QB Items and BOMs" & vbNewLine)
			End If

		Catch ex As Exception
			MsgBox(ex.Message)
		End Try
	End Sub

#Region "IMPORT SMT BOM"
	Private Sub GoThroughALPHABoards(ByRef FolderDirectiory As String)
		'Go through each 'Board' Folder.
		For Each dir As String In Directory.GetDirectories(FolderDirectiory)
			GoThroughALPHARevisions(dir)
		Next
	End Sub

	Private Sub GoThroughALPHARevisions(ByRef FolderDirectiory As String)
		'Go through each 'Revision' Folder.
		For Each dir As String In Directory.GetDirectories(FolderDirectiory)
			If dir.ToLower.Contains("oldreleases") Then
				Continue For
			End If
			GoThroughALPHAFiles(dir)
		Next
	End Sub

	Private Sub GoThroughALPHAFiles(ByRef FolderDirectory As String)
		'Grab all of the files in the folder and check to see if we have any
		Dim files() As String = Directory.GetFiles(FolderDirectory, ALPHA_EXE)
		If files.Length = 0 Then
			UpdateRTB("    No Files for: " & FolderDirectory & vbNewLine)
			ALPHAminor = True
			Return
		End If

		'Go through each file.
		For Each dataFile As String In Directory.GetFiles(FolderDirectory, ALPHA_EXE)
			Dim errorMessage As New List(Of String)
			Dim fileInformation As New FileInfo(dataFile)
			Dim fileNameParsed() As String = fileInformation.Name.Split(PERIOD_DILIMITER)

			'Check to see if the File is larger than 0 in size.
			If fileInformation.Length = 0 Then
				ALPHAminor = True
				UpdateRTB("    FAILED to import " & fileInformation.Name & " File size is 0" & vbNewLine)
				Continue For
			End If

			'check to see if the file has the right number of 'parts' in the name. If not, add error and move on to the next file.
			If fileNameParsed.Length < 4 Then
				ALPHAmajor = True
				UpdateRTB("    FAILED to import " & fileInformation.Name & " Check name format" & vbNewLine)
				Continue For
			End If

			'We have enough 'parts' so we can now parse the fileName.
			Dim fileName As String = fileNameParsed(INDEX_BOARD) & "." & fileNameParsed(INDEX_REVISION1) & "." & fileNameParsed(INDEX_REVISION2) & "." & fileNameParsed(INDEX_OPTION) & "."

			If ImportALPHAFile(dataFile, fileName, errorMessage) = False Then
				ALPHAmajor = True
				UpdateRTB("    FAILED to import " & fileName & vbNewLine)
				For Each errorItem In errorMessage
					UpdateRTB("        " & errorItem & vbNewLine)
				Next
			Else
				UpdateRTB("    Imported: " & fileName & vbNewLine)
			End If
			
		Next
	End Sub

	Private Function ImportALPHAFile(ByRef dataFile As String, ByRef fileName As String, ByRef errorMessage As List(Of String)) As Boolean
		Dim errorsExist As Boolean = False

		'F1 Parse
		Dim name As String = ""

		'F8 Parse
		Dim xPosition As Integer = Nothing
		Dim yPosition As Integer = Nothing
		Dim angle As Integer = Nothing
		Dim group As String = ""
		Dim mountSkip As String = ""
		Dim despenseSkip As String = ""
		Dim component As String = ""

		'F9 Parse
		Dim referenceDesignator As String = ""

		Dim parsedName As String() = fileName.Split(PERIOD_DILIMITER)

		'Start our transaction. Must assign both transaction object and connection to the command object for a pending local transaction.
		Dim transaction As SqlTransaction = Nothing
		transaction = myConn.BeginTransaction("SMT Transaction")
		myCmd.Connection = myConn
		myCmd.Transaction = transaction

		Try
			Using myParser As New TextFieldParser(dataFile)
				myParser.TextFieldType = FieldType.Delimited
				myParser.SetDelimiters(SPACE_DILIMITER)

				Dim currentRow As String()

				'Read the first line so we can ignore it.
				currentRow = myParser.ReadFields()

				While Not myParser.EndOfData
					currentRow = myParser.ReadFields()

					Select Case currentRow(ALHPA_CODE_INDEX)
						Case "F1"
							'F1 ETSMBWB.Rev1.1.AD.
							'This line code should be the first line in the file and only happen once.

							name = currentRow(ALPHA_NAME_INDEX)
							If String.Compare(name, fileName, True) <> 0 Then
								errorMessage.Add("File name [" & fileName & "] and Name[" & name & "] did not match. See Line F1.")
								errorsExist = True
							End If
						Case "F8"
							'F8 74127 107077 0 0 N N SMT102C25NP00603
							'F9 C1

							xPosition = currentRow(ALHPA_X_POSITION_INDEX)
							yPosition = currentRow(ALHPA_Y_POSITION_INDEX)
							angle = currentRow(ALHPA_ANGLE_INDEX)
							group = currentRow(ALHPA_GROUP_INDEX)
							mountSkip = currentRow(ALHPA_MOUNT_SKIP_INDEX)
							despenseSkip = currentRow(ALHPA_DESPENSE_SKIP_INDEX)
							component = currentRow(ALHPA_COMPONENT_INDEX)

							'The next line should always be an F9 so we should be able to just read the next line.
							currentRow = myParser.ReadFields()
							referenceDesignator = currentRow(ALHPA_REFERENCE_DESIGNATOR_INDEX)

							'Create our new database entry.
							If errorsExist = True Then
								Continue While
							End If
							myCmd.CommandText = "INSERT INTO dbo." & TABLE_ALPHABOM & "([" & DB_HEADER_REF_DES & "], [" & DB_HEADER_ITEM_NUMBER & "], [" & DB_HEADER_PROCESS & "], [" & DB_HEADER_BOARD_NAME & "], [" & DB_HEADER_POS_X & "], [" & DB_HEADER_POS_Y & "], [" & DB_HEADER_ANGLE & "], [" & DB_HEADER_GROUP & "], [" & DB_HEADER_MOUNT_SKIP & "], [" & DB_HEADER_DESPENSE_SKIP & "]) " &
												"VALUES('" & referenceDesignator & "', '" & component & "', '" & PROCESS_SMT & "', '" & name & "', " & xPosition & ", " & yPosition & ", " & angle & ", '" & group & "', '" & mountSkip & "', '" & despenseSkip & "')"
							myCmd.ExecuteNonQuery()
						Case "#"
							Exit While
					End Select
				End While
			End Using

			If errorsExist = True Then
				sqlapi.RollBack(transaction, errorMessage)
				Return False
			End If

			transaction.Commit()
			Return True
		Catch ex As Exception
			errorMessage.Add(ex.Message)

			'If we have started our transaction, we need to close it correctly.
			If Not transaction Is Nothing Then
				sqlapi.RollBack(transaction, errorMessage)
			End If
			Return False
		End Try
	End Function
#End Region

#Region "IMPORT SMT ITEMS"
	Private Sub GoThroughALPHAItems()
		ImportSMTmagazines()

		ImportSMTItems()

		ImportSMTPackages()
	End Sub

	Public Sub ImportSMTMagazines()
		UpdateRTB("Starting to import SMT Magazines [" & AlphaItemslocation & "]" & vbNewLine)

		' check for only one file in the location
		Dim fileEntries As String() = Directory.GetFiles(AlphaItemslocation & "Alex's Test\", "*.mag")
		Dim results As String = ""

		If fileEntries.Count = 0 Then
			UpdateRTB("    No .mag file found. Skipping Import" & vbNewLine & vbNewLine)

			ALPHAitemsMajor = True
			Return
		End If

		Dim isFirst As Boolean = True

		' go through each mag file that we have
		For Each filepath In fileEntries
			Dim fileinfo As New FileInfo(filepath)
			UpdateRTB("    " & fileinfo.Name & " ")

			ImportSMTMagazine(filepath, isFirst, results, True)

			If results.Length <> 0 Then
				UpdateRTB("Error:" & vbNewLine & "    " & results)
			Else
				UpdateRTB("O.K." & vbNewLine)
			End If

			isFirst = False
		Next

		UpdateRTB("Finished importing SMT Magazines" & vbNewLine & vbNewLine)
	End Sub

	Public Function ImportSMTMagazine(ByVal filePath As String, ByRef isFirst As Boolean, ByRef result As String, Optional updateUI As Boolean = False) As Boolean
		Dim line As String = ""
		Dim prefix As String = ""

		'M10 Parse
		Dim name As String = ""
		Dim serialNumber As String = ""

		'M15 Parse
		Dim machinenumber As String = ""
		Dim machineString As String = ""
		Dim slotNumber As Integer = 0

		'M21 Parse
		Dim feederNumber As Integer = 0
		Dim angle As Integer = 0
		Dim quantity As Integer = 0
		Dim stockNumber As String = ""

		Dim transaction As SqlTransaction = Nothing

		Dim isDeleted As Boolean = False

		transaction = myConn.BeginTransaction("Magazine Transaction")
		myCmd.Connection = myConn
		myCmd.Transaction = transaction

		Dim totalLines As Integer = File.ReadAllLines(filePath).Length
		Dim sr As New StreamReader(filePath)
		

		Try
			Dim itemcount As Integer = 0
			Dim periodCount As Integer = 0
			While Not sr.EndOfStream
				itemcount += 1

				If itemcount >= (totalLines \ 20) and periodCount <> 20 Then
					'The purpose of this is to let the user know that we are still adding items and the progam is not crashed.
					UpdateRTB(". ")
					itemcount = 0
					periodCount +=1
				End If

				line = sr.ReadLine
				If line.Length <> 1 Then
					prefix = line.Substring(0, line.IndexOf(" "))
				Else
					prefix = line
				End If

				'Crazy math. The way that this file is set up prevents us from spliting each line perfectly. There are items that have '"' with spaces
				'which make parsing the file difficult. To get around this, we use 'GetNthIndex' to get the start index of the # of a given symbol.
				'		M10 48 132451 -84340 "AGILIS 8-1" "" 0 0 0 Y
				'AGILIS 8-1 is the whole name so we need the index of the 1st '"' + 1 to the index of the 2nd '"'.
				'																+ 1 to not include the beginning '"'
				'The length is then calculated index of the 2nd '"' - index of the 1st '"' - 1
				'																- 1 to not include the ending '"'
				'Same rule applies to the spaces. 
				Select Case prefix
					Case "M10"
						'M10 48 132451 -84340 "AGILIS 8-1" "" 0 0 0 Y

						name = line.Substring(GetNthIndex(line, """", 1) + 1, GetNthIndex(line, """", 2) - GetNthIndex(line, """", 1) - 1)
						serialNumber = line.Substring(GetNthIndex(line, " ", 2) + 1, GetNthIndex(line, " ", 3) - GetNthIndex(line, " ", 2) - 1)
						isDeleted = False

					Case "M15"
						'M15 0 -1

						machinenumber = CInt(line.Substring(GetNthIndex(line, " ", 1) + 1, GetNthIndex(line, " ", 2) - GetNthIndex(line, " ", 1) - 1))
						slotNumber = CInt(line.Substring(GetNthIndex(line, " ", 2) + 1)) + 1

						Select machinenumber
						    Case NOTLOADED_NUM
								machineString = NOTLOADED

							Case ALPHA_NUM
								machineString = ALPHA

							Case GAMMA_NUM
								machineString = GAMMA

						End Select
					Case "M21"
						'M21 1 false 8mmtape [90000 -1 "SMT 330R1206-5" ""]

						' check to see if we are not on the first file anymore.
						If isFirst = False And machinenumber = NOTLOADED_NUM Then
							' if we are not the first file, we only want to update what is loaded in a machine.
							Continue While
						End If

						If isFirst = False And isDeleted = False Then
							' if we are not on the first file and we have a machine number, we need to clear any existing data with the magazine
							myCmd.CommandText = "DELETE FROM " & TABLE_MAGAZINE_DATA & " WHERE " & DB_HEADER_NAME & " = '" & name & "'"
							myCmd.ExecuteNonQuery()
							isDeleted = True
						End If

						feederNumber = CInt(line.Substring(GetNthIndex(line, " ", 1) + 1, GetNthIndex(line, " ", 2) - GetNthIndex(line, " ", 1) - 1)) + 1
						Try
							angle = line.Substring(GetNthIndex(line, " ", 4) + 1, GetNthIndex(line, " ", 5) - GetNthIndex(line, " ", 4) - 1)
							quantity = line.Substring(GetNthIndex(line, " ", 5) + 1, GetNthIndex(line, " ", 6) - GetNthIndex(line, " ", 5) - 1)
							stockNumber = line.Substring(GetNthIndex(line, """", 1) + 1, GetNthIndex(line, """", 2) - GetNthIndex(line, """", 1) - 1)
							myCmd.CommandText = "INSERT INTO " & TABLE_MAGAZINE_DATA & " ([" & DB_HEADER_SERIAL_NUMBER & "], [" & DB_HEADER_NAME & "], [" & DB_HEADER_MACHINE_NUMBER & "], [" & DB_HEADER_SLOT_NUMBER & "], [" & DB_HEADER_FEEDER_NUMBER & "], [" & DB_HEADER_ANGLE & "], [" & DB_HEADER_QUANTITY & "], [" & DB_HEADER_ITEM_NUMBER & "]) " &
												"VALUES('" & serialNumber & "', '" & name & "', '" & machineString & "', " & slotNumber & ", " & feederNumber & ", " & angle & ", " & quantity & ", '" & stockNumber & "')"
							myCmd.ExecuteNonQuery()
						Catch ex As Exception
							'The only time that we make it here is if the line does not give us this [extra information]. They are not required for the machine file.
							myCmd.CommandText = "INSERT INTO " & TABLE_MAGAZINE_DATA & " ([" & DB_HEADER_SERIAL_NUMBER & "], [" & DB_HEADER_NAME & "], [" & DB_HEADER_MACHINE_NUMBER & "], [" & DB_HEADER_SLOT_NUMBER & "], [" & DB_HEADER_FEEDER_NUMBER & "]) " &
												"VALUES('" & serialNumber & "', '" & name & "', '" & machineString & "', " & slotNumber & ", " & feederNumber & ")"
							myCmd.ExecuteNonQuery()
						End Try
				End Select
			End While
			transaction.Commit()
		Catch ex As Exception
			sqlapi.RollBack(transaction, errorMessage:=New List(Of String))
			result = ex.Message
			Return False
		End Try

		Return True
	End Function

	Private Sub ImportSMTItems()
		UpdateRTB("Starting to import SMT Items [" & AlphaItemslocation & "]" & vbNewLine)

		' check for only one file in the location
		Dim fileEntries As String() = Directory.GetFiles(AlphaItemslocation & "Alex's Test\", "*.cmp")
		Dim results As String = ""

		If fileEntries.Count = 0 Then
			UpdateRTB("    No .cmp file found. Skipping Import" & vbNewLine & vbNewLine)

			ALPHAitemsMajor = True
			Return
		ElseIf 1 < fileEntries.Count then	
			UpdateRTB("    More than 1 .cmp file found. Skipping Import" & vbNewLine & vbNewLine)

			ALPHAitemsMajor = True
			Return
		End If

		ImportSMTItem(fileEntries(0), results)

		If results.Length <> 0 Then
			UpdateRTB(results & vbNewLine)
		End If

		UpdateRTB("Finished importing SMT Items" & vbNewLine & vbNewLine)
	End Sub

	Public Function ImportSMTItem(ByVal filePath As String, ByRef results As String) As Boolean
		Dim line As String = ""
		Dim prefix As String = ""

		'C00 Parse
		Dim stockNumber As String = ""

		'C10 Parse
		Dim defaultTapeAngle As String = ""

		'C11 Parse
		Dim defaultSteps As String = ""
		Dim finepitch As String = ""

		'C12 Parse
		Dim defaultStepLength As String = ""
		Dim defaultStepLengthTrim As String = ""

		'Generic Parse
		Dim value As String = ""

		Dim transaction As SqlTransaction = Nothing

		transaction = myConn.BeginTransaction("Component Transaction")
		myCmd.Connection = myConn
		myCmd.Transaction = transaction

		Dim totalLines As Integer = File.ReadAllLines(filePath).Length
		Dim sr As New StreamReader(filePath)

		UpdateRTB("    ")

		Try
			Dim itemcount As Integer = 0
			Dim periodCount As Integer = 0
			While Not sr.EndOfStream
				itemcount += 1

				If itemcount >= (totalLines \ 20) and periodCount <> 20 Then
					'The purpose of this is to let the user know that we are still adding items and the progam is not crashed.
					UpdateRTB(". ")
					itemcount = 0
					periodCount +=1
				End If
				
				line = sr.ReadLine

				If line.Length <> 1 Then
					prefix = line.Substring(0, line.IndexOf(" "))
				Else
					prefix = line
				End If

				'Crazy math. The way that this file is set up prevents us from spliting each line perfectly. There are items that have '"' with spaces
				'which make parsing the file difficult. To get around this, we use 'GetNthIndex' to get the start index of the # of a given symbol.
				'		M10 48 132451 -84340 "AGILIS 8-1" "" 0 0 0 Y
				'AGILIS 8-1 is the whole name so we need the index of the 1st '"' + 1 to the index of the 2nd '"'.
				'																+ 1 to not include the beginning '"'
				'The length is then calculated index of the 2nd '"' - index of the 1st '"' - 1
				'																- 1 to not include the ending '"'
				'Same rule applies to the spaces. 
				Select Case prefix
					Case "C00"
						'C00 ABRACON 24.000MHZ

						stockNumber = line.Substring(line.IndexOf(" ") + 1)
						myCmd.CommandText = "INSERT INTO " & TABLE_ALPHA_ITEMS & " ([" & DB_HEADER_ITEM_NUMBER & "]) VALUES('" & stockNumber & "')"
					Case "C01"
						'C01 ABM3B

						value = line.Substring(line.IndexOf(" ") + 1)
						myCmd.CommandText = "UPDATE " & TABLE_ALPHA_ITEMS & " SET [" & DB_HEADER_PACKAGE & "] = '" & value & "' WHERE [" & DB_HEADER_ITEM_NUMBER & "] = '" & stockNumber & "'"
					Case "C08"
						'C08 8mm

						value = line.Substring(line.IndexOf(" ") + 1)
						myCmd.CommandText = "UPDATE " & TABLE_ALPHA_ITEMS & " SET [" & DB_HEADER_FEEDER_TYPE & "] = '" & value & "' WHERE [" & DB_HEADER_ITEM_NUMBER & "] = '" & stockNumber & "'"
					Case "C081"
						'C081 TAPE_MAG

						value = line.Substring(line.IndexOf(" ") + 1)
						myCmd.CommandText = "UPDATE " & TABLE_ALPHA_ITEMS & " SET [" & DB_HEADER_MAGAZINE_TYPE & "] = '" & value & "' WHERE [" & DB_HEADER_ITEM_NUMBER & "] = '" & stockNumber & "'"
					Case "C10"
						'C10 90000 0

						defaultTapeAngle = line.Substring(GetNthIndex(line, " ", 1) + 1, GetNthIndex(line, " ", 2) - GetNthIndex(line, " ", 1) - 1)
						myCmd.CommandText = "UPDATE " & TABLE_ALPHA_ITEMS & " SET [" & DB_HEADER_DEFAULT_TAPE_ANGLE & "] = '" & defaultTapeAngle & "' WHERE [" & DB_HEADER_ITEM_NUMBER & "] = '" & stockNumber & "'"
					Case "C11"
						'

						defaultSteps = line.Substring(GetNthIndex(line, " ", 1) + 1, GetNthIndex(line, " ", 2) - GetNthIndex(line, " ", 1) - 1)
						finepitch = line.Substring(GetNthIndex(line, " ", 2) + 1)
						myCmd.CommandText = "UPDATE " & TABLE_ALPHA_ITEMS & " SET [" & DB_HEADER_DEFAULT_STEPS & "] = '" & defaultSteps & "', [" & DB_HEADER_FINEPITCH & "] = '" & finepitch & "' WHERE [" & DB_HEADER_ITEM_NUMBER & "] = '" & stockNumber & "'"
					Case "C12"
						'C12 4000 0

						defaultStepLength = line.Substring(GetNthIndex(line, " ", 1) + 1, GetNthIndex(line, " ", 2) - GetNthIndex(line, " ", 1) - 1)
						defaultStepLengthTrim = line.Substring(GetNthIndex(line, " ", 2) + 1)
						myCmd.CommandText = "UPDATE " & TABLE_ALPHA_ITEMS & " SET [" & DB_HEADER_DEFAULT_STEP_LENGTH & "] = '" & defaultStepLength & "', [" & DB_HEADER_DEFAULT_STEP_LENGTH_TRIM & "] = '" & defaultStepLengthTrim & "' WHERE [" & DB_HEADER_ITEM_NUMBER & "] = '" & stockNumber & "'"
				End Select

				myCmd.ExecuteNonQuery()
			End While

			transaction.Commit()
		Catch ex As Exception
			sqlapi.RollBack(transaction, errorMessage:=New List(Of String))
			results = ex.Message
			Return False
		End Try

		UpdateRTB("O.K." & vbNewLine)
		Return True
	End Function

	Private Sub ImportSMTPackages()
		UpdateRTB("Starting to import SMT Packages [" & AlphaItemslocation & "]" & vbNewLine)

		' check for only one file in the location
		Dim fileEntries As String() = Directory.GetFiles(AlphaItemslocation, "pck.gen")
		Dim results As String = ""

		If fileEntries.Count = 0 Then
			UpdateRTB("    No pck.gen file found. Skipping Import" & vbNewLine & vbNewLine)

			ALPHAitemsMajor = True
			Return
		ElseIf 1 < fileEntries.Count then	
			UpdateRTB("    More than 1 pck.gen file found. Skipping Import" & vbNewLine & vbNewLine)

			ALPHAitemsMajor = True
			Return
		End If

		ImportSMTPackage(fileEntries(0), results)

		If results.Length <> 0 Then
			UpdateRTB(results & vbNewLine)
		End If

		UpdateRTB("Finished importing SMT Packages" & vbNewLine & vbNewLine)
	End Sub

	Public Function ImportSMTPackage(ByVal filePath As String, ByRef results As String) As Boolean
		Dim line As String = ""
		Dim prefix As String = ""
		Dim parseSet As Integer = 1

		'P00 Parse
		Dim packageName As String = ""

		'P01 Parse
		Dim bodyLength As String = ""
		Dim bodyWidth As String = ""
		Dim overallLength As String = ""
		Dim overallWidth As String = ""
		Dim normalHeigth As String = ""
		Dim maxHeigth As String = ""
		Dim minHeigth As String = ""

		'P062-M  Parse
		Dim angle As String = ""
		Dim level As String = ""
		Dim position As String = ""
		Dim force As String = ""

		'P063  Parse
		Dim normalSize As String = ""
		Dim maxSize As String = ""
		Dim minSize As String = ""
		Dim verifyMechanical As String = ""

		Dim transaction As SqlTransaction = Nothing

		transaction = myConn.BeginTransaction("Package Transaction")
		myCmd.Connection = myConn
		myCmd.Transaction = transaction

		Dim totalLines As Integer = File.ReadAllLines(filePath).Length
		Dim sr As New StreamReader(filePath)
		UpdateRTB("    ")

		Try
			Dim itemcount As Integer = 0
			Dim periodCount As Integer = 0
			While Not sr.EndOfStream
				itemcount += 1

				If itemcount >= (totalLines \ 20) and periodCount <> 20 Then
					'The purpose of this is to let the user know that we are still adding items and the progam is not crashed.
					UpdateRTB(". ")
					itemcount = 0
					periodCount +=1
				End If
				
				line = sr.ReadLine
				If line.Length <> 1 Then
					prefix = line.Substring(0, line.IndexOf(" "))
				Else
					prefix = line
				End If

				Select Case prefix
					Case "P00"
						'P00 0402-05(a tiny bit small)

						packageName = line.Substring(line.IndexOf(" ") + 1)
						myCmd.CommandText = "INSERT INTO " & TABLE_ALPHA_PACKAGE & " ([" & DB_HEADER_PACKAGE_NAME & "]) VALUES('" & packageName & "')"
						myCmd.ExecuteNonQuery()

						'Reset our parse set to 1.
						parseSet = 1
					Case "P01"
						'P01 900 400 900 400 500 600 400

						bodyLength = line.Substring(GetNthIndex(line, " ", 1) + 1, GetNthIndex(line, " ", 2) - GetNthIndex(line, " ", 1) - 1) / 1000
						bodyWidth = line.Substring(GetNthIndex(line, " ", 2) + 1, GetNthIndex(line, " ", 3) - GetNthIndex(line, " ", 2) - 1) / 1000
						overallLength = line.Substring(GetNthIndex(line, " ", 3) + 1, GetNthIndex(line, " ", 4) - GetNthIndex(line, " ", 3) - 1) / 1000
						overallWidth = line.Substring(GetNthIndex(line, " ", 4) + 1, GetNthIndex(line, " ", 5) - GetNthIndex(line, " ", 4) - 1) / 1000
						normalHeigth = line.Substring(GetNthIndex(line, " ", 5) + 1, GetNthIndex(line, " ", 6) - GetNthIndex(line, " ", 5) - 1) / 1000
						maxHeigth = line.Substring(GetNthIndex(line, " ", 6) + 1, GetNthIndex(line, " ", 7) - GetNthIndex(line, " ", 6) - 1) / 1000
						minHeigth = line.Substring(line.LastIndexOf(" ") + 1) / 1000

						myCmd.CommandText = "UPDATE " & TABLE_ALPHA_PACKAGE & " SET [" & DB_HEADER_BODY_LENGTH & "] = " & bodyLength & ", [" & DB_HEADER_BODY_WIDTH & "] = " & bodyWidth & ", [" & DB_HEADER_OVERALL_LENGTH & "] = " & overallLength & ", [" & DB_HEADER_OVERALL_WIDTH & "] = " & overallWidth & ", [" & DB_HEADER_NORMAL_HEIGHT & "] = " & normalHeigth & ", [" & DB_HEADER_MAX_HEIGHT & "] = " & maxHeigth & ", [" & DB_HEADER_MIN_HEIGHT & "] = " & minHeigth & " WHERE [" & DB_HEADER_PACKAGE_NAME & "] = '" & packageName & "'"
						myCmd.ExecuteNonQuery()
					Case "P062-M"
						'P062-M 90000 0 POS_UPPER MIDDLE_FORCE

						angle = line.Substring(GetNthIndex(line, " ", 1) + 1, GetNthIndex(line, " ", 2) - GetNthIndex(line, " ", 1) - 1) / 1000
						level = line.Substring(GetNthIndex(line, " ", 2) + 1, GetNthIndex(line, " ", 3) - GetNthIndex(line, " ", 2) - 1)
						position = line.Substring(GetNthIndex(line, " ", 3) + 1, GetNthIndex(line, " ", 4) - GetNthIndex(line, " ", 3) - 1)
						force = line.Substring(line.LastIndexOf(" ") + 1)

						If parseSet = 1 Then
							myCmd.CommandText = "UPDATE " & TABLE_ALPHA_PACKAGE & " SET [" & DB_HEADER_ANGLE_1 & "] = '" & angle & "', [" & DB_HEADER_LEVEL_1 & "] = '" & level & "', [" & DB_HEADER_POSITION_1 & "] = '" & position & "', [" & DB_HEADER_FORCE_1 & "] = '" & force & "' WHERE [" & DB_HEADER_PACKAGE_NAME & "] = '" & packageName & "'"
						ElseIf parseSet = 2 Then
							myCmd.CommandText = "UPDATE " & TABLE_ALPHA_PACKAGE & " SET [" & DB_HEADER_ANGLE_2 & "] = '" & angle & "', [" & DB_HEADER_LEVEL_2 & "] = '" & level & "', [" & DB_HEADER_POSITION_2 & "] = '" & position & "', [" & DB_HEADER_FORCE_2 & "] = '" & force & "' WHERE [" & DB_HEADER_PACKAGE_NAME & "] = '" & packageName & "'"
						End If

						myCmd.ExecuteNonQuery()
					Case "P063"
						'P063 1000 1300 700 true

						normalSize = line.Substring(GetNthIndex(line, " ", 1) + 1, GetNthIndex(line, " ", 2) - GetNthIndex(line, " ", 1) - 1) / 1000
						maxSize = line.Substring(GetNthIndex(line, " ", 2) + 1, GetNthIndex(line, " ", 3) - GetNthIndex(line, " ", 2) - 1) / 1000
						minSize = line.Substring(GetNthIndex(line, " ", 3) + 1, GetNthIndex(line, " ", 4) - GetNthIndex(line, " ", 3) - 1) / 1000
						verifyMechanical = line.Substring(line.LastIndexOf(" ") + 1)

						If parseSet = 1 Then
							myCmd.CommandText = "UPDATE " & TABLE_ALPHA_PACKAGE & " SET [" & DB_HEADER_NORMAL_SIZE_1 & "] = " & normalSize & ", [" & DB_HEADER_MAX_SIZE_1 & "] = " & maxSize & ", [" & DB_HEADER_MIN_SIZE_1 & "] = " & minSize & ", [" & DB_HEADER_VERIFY_MECHANICAL_1 & "] = '" & verifyMechanical & "' WHERE [" & DB_HEADER_PACKAGE_NAME & "] = '" & packageName & "'"

							'Increase our Parse set to 2 incase we have a second pass.
							parseSet = 2
						ElseIf parseSet = 2 Then
							myCmd.CommandText = "UPDATE " & TABLE_ALPHA_PACKAGE & " SET [" & DB_HEADER_NORMAL_SIZE_2 & "] = " & normalSize & ", [" & DB_HEADER_MAX_SIZE_2 & "] = " & maxSize & ", [" & DB_HEADER_MIN_SIZE_2 & "] = " & minSize & ", [" & DB_HEADER_VERIFY_MECHANICAL_2 & "] = '" & verifyMechanical & "' WHERE [" & DB_HEADER_PACKAGE_NAME & "] = '" & packageName & "'"
						End If

						myCmd.ExecuteNonQuery()
				End Select
			End While

			transaction.Commit()
		Catch ex As Exception
			sqlapi.RollBack(transaction, errorMessage:=New List(Of String))
			results = ex.Message
			Return False
		End Try

		UpdateRTB("O.K." & vbNewLine)
		Return True
	End Function

#End Region

#Region "IMPORT PCAD BOM"
	Private Sub GoThroughImportPCADBoards(ByRef FolderDirectiory As String)
		'Go through each 'Board' Folder.
		For Each dir As String In Directory.GetDirectories(FolderDirectiory)
			GoThroughPCADRevisions(dir)
		Next
	End Sub

	Private Sub GoThroughPCADRevisions(ByRef FolderDirectiory As String)
		'Go through each 'Revision' Folder.
		For Each dir As String In Directory.GetDirectories(FolderDirectiory)
			If dir.ToLower.Contains("oldreleases") Then
				Continue For
			End If

			If GoThroughPCADFiles(dir) = True Then
				Dim missingFiles As New List(Of String)

				'Check to see if this revision folder is 'Release Ready'.
				If CheckforRelease(dir, missingFiles, needsPNP) = False Then
					UpdateRTB("        NOT RELEASE READY" & vbNewLine)
					For Each item In missingFiles
						UpdateRTB("            " & item & vbNewLine)
					Next
					PCADminor = True
				Else
					UpdateRTB("        Release Ready" & vbNewLine)
				End If
			End If
		Next
	End Sub

	Private Function GoThroughPCADFiles(ByRef dir As String) As Boolean
		'Check to see if we have any files in the directory to import.
		Dim files() As String = Directory.GetFiles(dir, PCAD_EXE)
		If files.Length = 0 Then
			UpdateRTB("    No Files for: " & dir & vbNewLine)
			PCADminor = True
			Return False
		End If

		'Go through each file.
		For Each dataFile As String In Directory.GetFiles(dir, PCAD_EXE)
			Dim errorMessage As New List(Of String)
			Dim fileInformation As New FileInfo(dataFile)
			Dim fileNameParsed() As String = fileInformation.Name.Split(PERIOD_DILIMITER)

			'check to see if the file has the right number of 'parts' in the name. If not, add error and move on to the next file.
			If fileNameParsed.Length < 4 Then
				PCADmajor = True
				UpdateRTB("    FAILED to import " & fileInformation.Name & " Check name format" & vbNewLine)
				Continue For
			End If

			'We have enough 'parts' so we can now parse the fileName.
			Dim fileName As String = fileNameParsed(INDEX_BOARD) & "." & fileNameParsed(INDEX_REVISION1) & "." & fileNameParsed(INDEX_REVISION2) & "." & fileNameParsed(INDEX_OPTION) & "."

			If ImportPCADFile(dataFile, fileName, errorMessage) = False Then
				PCADmajor = True
				UpdateRTB("    FAILED to import " & fileName & vbNewLine)
				For Each errorItem In errorMessage
					UpdateRTB("        " & errorItem & vbNewLine)
				Next
			Else
				If errorMessage.Count <> 0 Then
					UpdateRTB("    " & fileName & " was imported with minor errors:" & vbNewLine)
					For Each errorItem In errorMessage
						UpdateRTB("        " & errorItem & vbNewLine)
					Next
				Else
					UpdateRTB("    Imported: " & fileName & vbNewLine)
				End If
			End If
		Next

		Return True
	End Function

	Private Function ImportPCADFile(ByRef dataFile As String, ByVal fileName As String, ByRef errorMessage As List(Of String)) As Boolean
		needsPNP = False

		Dim errorsExist As Boolean = False
		Dim originalName As String = fileName
		Dim hasOptionDescription As Boolean = False
		Dim hasZF As Boolean = False
		Dim hasSMT As Boolean = False
		Dim fileNameParsed() As String = fileName.Split(PERIOD_DILIMITER)
		fileName = fileNameParsed(INDEX_BOARD) & "." & fileNameParsed(INDEX_REVISION1) & "." & fileNameParsed(INDEX_REVISION2) & ".."

		'Parse Values
		Dim name As String = ""
		Dim referenceDesignator As String = ""
		Dim value As String = ""
		Dim vendor As String = ""
		Dim partNumber As String = ""
		Dim stockNumber As String = ""
		Dim process As String = ""
		Dim prefix As String = ""

		'Optional Values
		Dim optionValue As String = ""

		'Indexs
		Dim INDEX_refdes As Integer = -1
		Dim INDEX_value As Integer = -1
		Dim INDEX_vendor As Integer = -1
		Dim INDEX_partNumber As Integer = -1
		Dim INDEX_stockNumber As Integer = -1
		Dim INDEX_process As Integer = -1

		'Optional Indexs
		Dim INDEX_options As Integer = -1
		Dim INDEX_swap As Integer = -1

		'Start our transaction. Must assign both transaction object and connection to the command object for a pending local transaction.
		Dim transaction As SqlTransaction = Nothing
		transaction = myConn.BeginTransaction("PCAD Transaction")
		myCmd.Connection = myConn
		myCmd.Transaction = transaction

		Try
			Using myParser As New TextFieldParser(dataFile)
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

				'Check to make sure the file has all of the required headers.
				If INDEX_refdes = -1 Then
					errorMessage.Add("No Reference Designator field")
					errorsExist = True
				End If
				If INDEX_value = -1 Then
					errorMessage.Add("No Value field")
					errorsExist = True
				End If
				If INDEX_vendor = -1 Then
					errorMessage.Add("No Vendor field")
					errorsExist = True
				End If
				If INDEX_partNumber = -1 Then
					errorMessage.Add("No Part Number field")
					errorsExist = True
				End If
				If INDEX_stockNumber = -1 Then
					errorMessage.Add("No Stock Number field")
					errorsExist = True
				End If
				If INDEX_process = -1 Then
					errorMessage.Add("No Process field")
					errorsExist = True
				End If

				'If we have syntax errors at this point, do not read the file and continue.
				If errorsExist = True Then
					sqlapi.RollBack(transaction, errorMessage)
					Return False
				End If

				While Not myParser.EndOfData
					'Reset everything to a blank state.
					name = ""
					referenceDesignator = ""
					value = ""
					vendor = ""
					partNumber = ""
					stockNumber = ""
					process = ""
					prefix = ""

					currentRow = myParser.ReadFields()

					'Check to see if we have something to read. If not, continue to the next line.
					If currentRow(0).Length = 0 Then
						Continue While
					End If

					'- - - Parse Reference Designator - - -

					referenceDesignator = currentRow(INDEX_refdes)

					'Check to see if we can swap the reference designator.
					If INDEX_swap <> -1 Then
						If currentRow(INDEX_swap).Length <> 0 Then
							referenceDesignator = currentRow(INDEX_swap)
						End If
					End If

					'A 'ZF' is only used when we have a BOM that does not have any SMT parts. This does not need to
					'be added to the database.
					If referenceDesignator.Contains("ZF") Then
						hasZF = True
						If currentRow(INDEX_value).Contains("NO SMT") = False Then
							'We have a ZF that is not supposed to be part of the BOM.
							errorMessage.Add("Reference ZF present when not needed.")
							hasZF = False
						End If
						Continue While
					End If

					'- - - Parse Process - - -

					process = currentRow(INDEX_process).ToUpper

					'Check to see if we have any curly braces or is empty. - denotes a default value.
					If process.Length = 0 Or process.Contains("{"c) = True Or process.Contains("}"c) = True Then
						'Check to see if this is a 'ZD'. These ones are allowed to ignore this rule.
						If referenceDesignator.Contains(REFERENCE_DESIGNATOR_OPTION) = False Then
							errorMessage.Add(referenceDesignator & " Process is empty")
							errorsExist = True
						End If
					End If

					'Check to see if we have a SMT part. If we do then we need to have a PNP file.
					If String.Compare(process, PROCESS_SMT, True) = 0 Then
						hasSMT = True
					End If

					'Check to see if we are using one of the valid process' that we have defined.
					If String.Compare(process, PROCESS_SMT, True) <> 0 And
						String.Compare(process, PROCESS_SMTBOTTOM, True) <> 0 And
						String.Compare(process, PROCESS_HANDFLOW, True) <> 0 And
						String.Compare(process, PROCESS_POSTASSEMBLY, True) <> 0 And
						String.Compare(process, PROCESS_PCBBOARD, True) <> 0 And
						String.Compare(process, PROCESS_SMTHAND, True) <> 0 And
						String.Compare(process, PROCESS_BAS, True) <> 0 Then
						'Check to see if the process is 'Not Used'. If yes then continue on to the next part.
						If String.Compare(process, PROCESS_NOTUSED, True) = 0 Then
							Continue While
						Else
							'Check to see if we are working with a 'ZD'.
							If referenceDesignator.Contains(REFERENCE_DESIGNATOR_OPTION) = False Then
								errorMessage.Add(referenceDesignator & " Process [" & process & "] does not follow the naming convention [See Naming Convention Document]")
								errorsExist = True
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

							'Next, check to see if our prefix is part of the Option Description.
							If referenceDesignator.Contains(REFERENCE_DESIGNATOR_OPTION) = True Then
								'Check to make sure the values match So we know that we have a description of the option.
								If String.Compare(fileNameParsed(INDEX_OPTION), optionValue) <> 0 Then
									Continue While
								Else
									hasOptionDescription = True
								End If
							Else
								'Check each letter of the option feild to see if the file calls for it.
								For index = 0 To optionValue.Length - 1
									If fileNameParsed(INDEX_OPTION).Contains(optionValue(index)) = True Then
										include = True
										Exit For
									End If
								Next

								'Finally check to see if we need to include this part or not.
								If include = False Then
									Continue While
								End If
							End If
						End If
					End If

					'- - - Parse Value - - -

					value = currentRow(INDEX_value)

					'Replace any single qoutes with double single quotes for SQL Format.
					If value.Contains("'"c) = True Then
						value = value.Replace("'", "''")
					End If

					'Check to see if we have any curly braces -denotes a default value.
					If value.Length = 0 Or value.Contains("{"c) = True Or value.Contains("}"c) = True Then
						errorMessage.Add(referenceDesignator & " Value is empty")
						errorsExist = True
					End If

					'- - - Parse Vendor - - -

					vendor = currentRow(INDEX_vendor)

					'Check to see if we have any curly braces -denotes a default value.
					If vendor.Length = 0 Or vendor.Contains("{"c) = True Or vendor.Contains("}"c) = True Then
						'Check to see if this is a 'ZD'.
						If referenceDesignator.Contains(REFERENCE_DESIGNATOR_OPTION) = False Then
							errorMessage.Add(referenceDesignator & " Vendor is empty")
							errorsExist = True
						End If
					End If

					'- - - Parse Part Number - - -

					partNumber = currentRow(INDEX_partNumber)

					'Check to see if we have any curly braces -denotes a default value.
					If partNumber.Length = 0 Or partNumber.Contains("{"c) = True Or partNumber.Contains("}"c) = True Then
						'Check to see if this is a 'ZD'.
						If referenceDesignator.Contains(REFERENCE_DESIGNATOR_OPTION) = False Then
							errorMessage.Add(referenceDesignator & " Part Number is empty")
							errorsExist = True
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
						If referenceDesignator.Contains(REFERENCE_DESIGNATOR_OPTION) = False And referenceDesignator.Contains(REFERENCE_DESIGNATOR_FIDUCIAL) = False Then
							errorMessage.Add(referenceDesignator & " Stock Number not in the correct format.")
							errorsExist = True
						End If
					End If

					'- - - Parse the Prefix. - - -

					'Check to see if our stock number has a colon as part of the name.
					If currentRow(INDEX_stockNumber).Contains(":") = True Then
						prefix = currentRow(INDEX_stockNumber).Substring(0, currentRow(INDEX_stockNumber).IndexOf(":"))

						'Check to see if our prefix = 'PCB'.
						If String.Compare(prefix, PREFIX_PCB, True) = 0 Then

							'Check to make sure the 'PCB' Stock Name and file name match.
							name = stockNumber.Substring(stockNumber.IndexOf("-") + 1)
							If String.Compare(name, fileName, True) <> 0 Then
								errorMessage.Add(referenceDesignator & " File name [" & fileName & "] and PCB Stock Name [" & name & "] did not match")
								errorsExist = True
							End If

							'Check to make sure the 'PCB' Part Name and file name match.
							name = partNumber.Substring(partNumber.IndexOf("-") + 1)
							If String.Compare(name, fileName, True) <> 0 Then
								errorMessage.Add(referenceDesignator & " File name [" & fileName & "] and PCB Part Name [" & name & "] did not match")
								errorsExist = True
							End If
						End If
					Else
						'Check to see if we are working with a 'ZD'.
						If referenceDesignator.Contains(REFERENCE_DESIGNATOR_OPTION) = False Then
							errorMessage.Add(referenceDesignator & " Stock Number does not have a prefix")
							errorsExist = True
						End If
					End If

					'If we have had any errors at this point, do not add them to the database. They will all be rolled back.
					If errorsExist = True Then
						Continue While
					End If

					myCmd.CommandText = "INSERT INTO dbo." & TABLE_PCADBOM & "([" & DB_HEADER_REF_DES & "], [" & DB_HEADER_DESCRIPTION & "], [" & DB_HEADER_VENDOR & "], [" & DB_HEADER_MPN & "], [" & DB_HEADER_ITEM_NUMBER & "], [" & DB_HEADER_PROCESS & "], [" & DB_HEADER_ITEM_PREFIX & "], [" & DB_HEADER_BOARD_NAME & "]) " &
										"VALUES('" & referenceDesignator & "', '" & value & "', '" & vendor & "', '" & partNumber & "', '" & stockNumber & "', '" & process & "', '" & prefix & "', '" & originalName & "')"
					myCmd.ExecuteNonQuery()
				End While
			End Using

			'If we have options in our filename, check to see that we have a 'ZD'.
			If fileNameParsed(INDEX_OPTION).Length <> 0 And hasOptionDescription = False Then
				errorMessage.Add("This board has options but does not have a ZD Reference Designator to descibe option '" & fileNameParsed(INDEX_OPTION) & "'")
				errorsExist = True
			End If

			If hasSMT = True Then
				needsPNP = True
			End If

			If hasZF = True Then
				needsPNP = False
			End If

			'Check to see if we had any errors.
			If errorsExist = True Then
				sqlapi.RollBack(transaction, errorMessage)
				Return False
			End If

			transaction.Commit()

			'Finally, check to see if we ahve any duplicate Reference Designators.
			myCmd.CommandText = "SELECT [" & DB_HEADER_REF_DES & "], COUNT(*) AS CountOf FROM " & TABLE_PCADBOM & " WHERE [" & DB_HEADER_BOARD_NAME & "] = '" & originalName & "' GROUP BY [" & DB_HEADER_REF_DES & "] HAVING COUNT(*) > 1 ORDER BY CountOf DESC"
			Dim myReader As SqlDataReader
			myReader = myCmd.ExecuteReader

			If myReader.HasRows Then
				While myReader.Read
					errorMessage.Add(myReader(DB_HEADER_REF_DES) & " is duplicated")
					PCADminor = True
				End While
			End If
			myReader.Close()
			Return True
		Catch ex As Exception
			errorMessage.Add(ex.Message)
			If Not transaction Is Nothing Then
				sqlapi.RollBack(transaction, errorMessage)
			End If
			Return False
		End Try
	End Function

	Private Function PNP_Check(ByRef filePath As String, ByRef errorMessage As List(Of String)) As Boolean
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

		'variables used to make sure we have the important information.
		Dim br_check As Integer = -1
		Dim fn_check As Integer = -1
		Dim fu1_check As Integer = -1
		Dim fu2_check As Integer = -1
		Dim fu3_check As Integer = -1

		Try
			Using myParser As New TextFieldParser(filePath)
				myParser.TextFieldType = FieldType.Delimited
				myParser.SetDelimiters(",")
				Dim currentRow As String()

				'Third row is the header row.
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
					errorMessage.Add("PNP File is missing the following Fields: " & missingFields)
					Return False
				End If

				'This count is used to find the 5 important items that are needed for release. This is all that we need to look for in this method.
				Dim count As Integer = 0

				While Not myParser.EndOfData
					If count = 5 Then
						Exit While
					End If

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
							If fu1_check <> 0 Then
								count += 1
							End If

							fu1_check = 0

							'Check for the correct format with the FU1.
							' BR90|FNNetbox.10
							If currentRow(INDEX_Value).Contains("|") = True Then
								Dim parsed() As String = currentRow(INDEX_Value).Split("|")
								For Each item In parsed
									Select Case item.Substring(0, 2).ToLower
										Case "br"
											If br_check <> 0 Then
												count += 1
											End If
											br_check = 0
										Case "fn"
											If fn_check <> 0 Then
												count += 1
											End If
											fn_check = 0
									End Select
								Next
								Continue While
							Else
								errorMessage.Add("FU1 '|' format")
								Return False
							End If
						Case "FU2"
							If fu2_check <> 0 Then
								count += 1
							End If
							fu2_check = 0
							Continue While
						Case "FU3"
							If fu3_check <> 0 Then
								count += 1
							End If
							fu3_check = 0
							Continue While
					End Select
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

				errorMessage.Add(infostring)
				Return False
			End If
		Catch ex As Exception
			errorMessage.Add("COULD NOT FIND PNP FILE AT: " & filePath)
			Return False
		End Try
		Return True
	End Function

	Private Function CheckforRelease(ByRef folderPath As String, ByRef errorMessage As List(Of String), ByRef PNPneeded As Boolean) As Boolean
		Try
			Dim allFiles() As String = Directory.GetFiles(folderPath)
			Dim bomFiles() As String = Directory.GetFiles(folderPath, PCAD_EXE)
			Dim pcbFiles() As String = Directory.GetFiles(folderPath, PCB_EXE)
			Dim pnpFiles() As String = Directory.GetFiles(folderPath, PNP_CSV_EXE)
			Dim schFiles() As String = Directory.GetFiles(folderPath, SCH_EXE)
			Dim schpdfFiles() As String = Directory.GetFiles(folderPath, SCH_PDF_EXE)

			'Get the list of options that are valid to check for
			If bomFiles.Length = 0 Then
				errorMessage.Add("Missing BOM.CSV")
			End If

			'Check for .PCB File
			If pcbFiles.Length = 0 Then
				errorMessage.Add("Missing .PCB")
			ElseIf pcbFiles.Length > 1 Then
				errorMessage.Add("Extra .PCB")
			End If

			'Check for .PNP.CSV File
			If PNPneeded = True Then
				If pnpFiles.Length = 0 Then
					errorMessage.Add("Missing .PNP.CSV")
				ElseIf pnpFiles.Length > 1 Then
					errorMessage.Add("Extra .PNP.CSV")
				Else
					PNP_Check(pnpFiles(0), errorMessage)
				End If
			End If

			'Check for .SCH File
			If schFiles.Length = 0 Then
				errorMessage.Add("Missing .SCH")
			ElseIf schFiles.Length > 1 Then
				errorMessage.Add("Extra .SCH")
			End If

			'Check for .SCH.PDF File
			If schpdfFiles.Length = 0 Then
				errorMessage.Add("Missing .SCH.PDF")
			ElseIf schpdfFiles.Length > 1 Then
				errorMessage.Add("Extra .SCH.PDF")
			End If

			'Check for .ZIP File
			Try
				Dim zipFiles() As String = Directory.GetFiles(folderPath & "\Released", ZIP_EXE)
				If zipFiles.Length = 0 Then
					errorMessage.Add("Missing .ZIP")
				ElseIf zipFiles.Length > 1 Then
					errorMessage.Add("Extra .ZIP")
				End If
			Catch ex As Exception
				errorMessage.Add("Missing Released Folder")
				errorMessage.Add("Missing .ZIP")
			End Try

			If errorMessage.Count = 0 Then
				Return True
			Else
				Return False
			End If
		Catch ex As Exception
			MsgBox(ex.Message)
			Return False
		End Try
	End Function
#End Region

#Region "IMPORT QB ITEMS"
	Public Sub ImportQBItems()
		Try
			Dim cmd As New OdbcCommand(ODBC_ItemQuery, _cn)

			Dim myTable As New DataTable()
			'Dim stopwatch As Stopwatch = Stopwatch.StartNew()
			myTable.Load(cmd.ExecuteReader())
			'stopwatch.Stop()
			'UpdateRTB("QB Items query = " & stopwatch.Elapsed.ToString & vbNewLine)

			'Start our transaction. Must assign both transaction object and connection to the command object for a pending local transaction.
			Dim transaction As SqlTransaction = Nothing
			transaction = myConn.BeginTransaction("QB Items Transaction")
			myCmd.Connection = myConn
			myCmd.Transaction = transaction

			Try
				For Each row As DataRow In myTable.Rows
					Dim type As String = row(DB_HEADER_TYPE).ToString
					Dim item_Number As String = row(DB_HEADER_ITEM_NUMBER).ToString
					Dim item_Prefix As String = row(DB_HEADER_ITEM_PREFIX).ToString
					Dim manufacturer2_Value As String = row(DB_HEADER_VENDOR2).ToString
					Dim partNumber2_Value As String = row(DB_HEADER_MPN2).ToString
					Dim manufacturer3_Value As String = row(DB_HEADER_VENDOR3).ToString
					Dim partNumber3_Value As String = row(DB_HEADER_MPN3).ToString
					Dim leadTime_Value As String = row(DB_HEADER_LEAD_TIME).ToString

					Dim description_Value As String = row(DB_HEADER_DESCRIPTION).ToString
					'Check for any single qoutes and replace them with double single qoutes.
					If description_Value.Contains("'"c) = True Then
						description_Value = description_Value.Replace("'", "''")
					End If
					'Check to see if we have a description or not.
					If description_Value.Length = 0 Then
						UpdateRTB("       " & item_Number & " does not have a Description." & vbNewLine)
						QBitemsMinor = True
					End If

					Dim quantity_Value As Decimal = Nothing
					'Check to see if we have a quantity.
					If row(DB_HEADER_QUANTITY).ToString.Length = 0 Then
						UpdateRTB("       " & item_Number & " does not have a Quantity." & vbNewLine)
						QBitemsMinor = True
						quantity_Value = 0.0
					Else
						quantity_Value = row(DB_HEADER_QUANTITY).ToString
					End If

					Dim cost_Value As Decimal = Nothing
					'Check to see if we have a cost.
					If row(DB_HEADER_COST).ToString.Length = 0 Then
						UpdateRTB("       " & item_Number & " does not have a Cost." & vbNewLine)
						QBitemsMinor = True
						cost_Value = 0.0
					Else
						cost_Value = row(DB_HEADER_COST).ToString
					End If

					Dim preferredVendor_Value As String = row(DB_HEADER_VENDOR).ToString
					'Check to see if we have a preferred Vendor.
					If preferredVendor_Value.Length = 0 Then
						'Check to see if we are using anything that does not need to have a Preferred Vendor.
						If type.ToLower.Contains("assembly") = False Then
							UpdateRTB("       " & item_Number & " does not have a Prefered Vendor." & vbNewLine)
							QBitemsMinor = True
						End If
					End If

					Dim partNumber_Value As String = row(DB_HEADER_MPN).ToString
					'Check to see if we have a Manufacturer Part Number.
					If partNumber_Value.Length = 0 Then
						'Check to see if we are using anything that does not need to have a Manufacturer Part Number.
						If type.ToLower.Contains("assembly") = False Then
							UpdateRTB("       " & item_Number & " does not have a Manufacturer Part Number." & vbNewLine)
							QBitemsMinor = True
						End If
					End If

					Dim minOrderQty_Value As String = row(DB_HEADER_MIN_ORDER_QTY).ToString
					'Check to see if we have commas in this field. Remove them
					If minOrderQty_Value.Contains(","c) = True Then
						minOrderQty_Value = minOrderQty_Value.Replace(",", "")
					End If

					Dim reOrderQty_Value As String = row(DB_HEADER_REORDER_QTY).ToString
					'Check to see if we have commas in this field. Remove them
					If reOrderQty_Value.Contains(","c) = True Then
						reOrderQty_Value = reOrderQty_Value.Replace(",", "")
					End If

					myCmd.CommandText = "INSERT INTO dbo." & TABLE_QB_ITEMS & "([" & DB_HEADER_TYPE & "], [" & DB_HEADER_ITEM_PREFIX & "], [" & DB_HEADER_ITEM_NUMBER & "], [" & DB_HEADER_DESCRIPTION & "], [" & DB_HEADER_QUANTITY & "], [" & DB_HEADER_COST & "], [" & DB_HEADER_VENDOR & "], [" & DB_HEADER_MPN & "], [" & DB_HEADER_VENDOR2 & "], [" & DB_HEADER_MPN2 & "], [" & DB_HEADER_VENDOR3 & "], [" & DB_HEADER_MPN3 & "], [" & DB_HEADER_LEAD_TIME & "], [" & DB_HEADER_MIN_ORDER_QTY & "], [" & DB_HEADER_REORDER_QTY & "]) " &
										"VALUES('" & type & "', '" & item_Prefix & "', '" & item_Number & "', '" & description_Value & "', " & quantity_Value & ", " & cost_Value & ", '" & preferredVendor_Value & "', '" & partNumber_Value & "', '" & manufacturer2_Value & "', '" & partNumber2_Value & "', '" & manufacturer3_Value & "', '" & partNumber3_Value & "', '" & leadTime_Value & "', '" & minOrderQty_Value & "', '" & reOrderQty_Value & "')"
					myCmd.ExecuteNonQuery()

					UpdateRTB("    Imported: " & item_Prefix & ":" & item_Number & vbNewLine)
				Next

				transaction.Commit()
			Catch ex As Exception
				UpdateRTB("       " & ex.Message & vbNewLine)
				If Not transaction Is Nothing Then
					sqlapi.RollBack(transaction, errorMessage:=New List(Of String))
				End If
				QBitemsMajor = True
			End Try
		Catch ex As Exception
			UpdateRTB("       " & ex.Message & vbNewLine)
			_cn.Close()
			QBitemsMajor = True
		End Try
	End Sub
#End Region

#Region "IMPORT QB BOM"
	Public Sub ImportQBBOMS()
		Try
			Dim syntaxerror As Boolean = False
			'First we need to check to make sure that we have the correct information in our database before we try to use it

			'Check for Assembly as a subset
			Dim cmd As New OdbcCommand("SELECT [ParentRefFullName] AS ""Name Prefix"",
										[Name] As ""Name"",
										[ItemInventoryAssemblyLnItemInventoryRefFullName] as ""Item Number""
										FROM itemInventoryAssemblyLine
										WHERE [IsActive] = 1
										AND [ItemInventoryAssemblyLnItemInventoryRefFullName] NOT LIKE '%:%'", _cn)

			Dim errorTable As New DataTable()
			errorTable.Load(cmd.ExecuteReader)

			If errorTable.Rows.Count <> 0 Then
				UpdateRTB("       " & "These Items all need to be part of a subset before we can import." & vbNewLine)

				For Each row In errorTable.Rows
					UpdateRTB("              " & row(DB_HEADER_ITEM_NUMBER).ToString & vbNewLine)
				Next

				syntaxerror = True
			End If

			'Check for Item as a sub item
			cmd.CommandText = "SELECT [ParentRefFullName] AS ""Name Prefix"",
								[Name] As ""Name"",
								[ItemInventoryAssemblyLnItemInventoryRefFullName] as ""Item Number""
								FROM itemInventoryAssemblyLine
								WHERE [IsActive] = 1
								AND [ItemInventoryAssemblyLnItemInventoryRefFullName] NOT LIKE '%:%'"

			errorTable = New DataTable()
			errorTable.Load(cmd.ExecuteReader)

			If errorTable.Rows.Count <> 0 Then
				UpdateRTB("       " & "These Assmblies all need to be part of a subset before we can import." & vbNewLine)

				For Each row In errorTable.Rows
					UpdateRTB("              " & row(DB_HEADER_NAME).ToString & vbNewLine)
				Next

				syntaxerror = True
			End If

			If syntaxerror = True Then
				skipQB = True
				QBmajor = True
				Return
			End If

			cmd.CommandText = ODBC_AssemblyQuery

			Dim myTable As New DataTable()
			'Dim stopwatch As Stopwatch = Stopwatch.StartNew()
			myTable.Load(cmd.ExecuteReader())
			'stopwatch.Stop()
			'UpdateRTB("QB BOMs query = " & stopwatch.Elapsed.ToString & vbNewLine)

			Dim QBItemList As New DataTable()
			myCmd.CommandText = "SELECT * FROM " & TABLE_QB_ITEMS
			QBItemList.Load(myCmd.ExecuteReader())

			'Start our transaction. Must assign both transaction object and connection to the command object for a pending local transaction.
			Dim transaction As SqlTransaction = Nothing
			transaction = myConn.BeginTransaction("QB BOMs Transaction")
			myCmd.Connection = myConn
			myCmd.Transaction = transaction
			Try
				Dim lastAssembly As String = ""

				For Each row As DataRow In myTable.Rows
					Dim item_Number As String = row(DB_HEADER_ITEM_NUMBER).ToString
					Dim name_Prefix As String = row(DB_HEADER_NAME_PREFIX).ToString
					Dim item_Prefix As String = row(DB_HEADER_ITEM_PREFIX).ToString
					Dim quantity_Needed As String = row(DB_HEADER_QUANTITY).ToString

					If item_Number.Length = 0 Then
						Continue For
					End If

					Dim name As String = ""
					If name_Prefix = PREFIX_FGS Then
						name = row(DB_HEADER_NAME).ToString
					Else
						name = row(DB_HEADER_NAME).ToString.Substring(row(DB_HEADER_NAME).ToString.IndexOf("-") + 1)
					End If

					'Figure out what process the part belongs in.
					Dim process As String = ""
					Select name_Prefix.ToUpper()
					    Case PREFIX_PCB
							process = "PCB BOARD"

						Case PREFIX_FGS
							process = "POST ASSEMBLY"

						Case PREFIX_BAS, PREFIX_BIS, PREFIX_DAS
							process = "HAND FLOW"

						Case PREFIX_SMA
							process = "SMT"
					End Select
					
					If item_Prefix.ToUpper() = PREFIX_PCB = True Then
						process = "PCB BOARD"
					ElseIf process = "" then
						UpdateRTB("       " & item_Number & " not associated with a process." & vbNewLine)
						QBminor = True
						process = "???"
					End If

					'Type and cost need to be found in the SQL database of the QB items.
					Dim drs() As DataRow = QBItemList.Select("[" & DB_HEADER_ITEM_NUMBER & "] = '" & item_Number & "'")

					If drs.Length = 0 Then
						UpdateRTB("       " & item_Number & " was not imported as an item under " & name_Prefix & "-" & name & ". Check for active status and type." & vbNewLine)
						QBmajor = True
						Continue For
					End If
					Dim type As String = drs(0)(DB_HEADER_TYPE)
					Dim cost As Decimal = drs(0)(DB_HEADER_COST)

					myCmd.CommandText = "INSERT INTO dbo." & TABLE_QBBOM & "([" & DB_HEADER_ITEM_NUMBER & "], [" & DB_HEADER_COST & "], [" & DB_HEADER_QUANTITY & "], [" & DB_HEADER_PROCESS & "], [" & DB_HEADER_NAME_PREFIX & "], [" & DB_HEADER_NAME & "], [" & DB_HEADER_TYPE & "], [" & DB_HEADER_ITEM_PREFIX & "]) " &
										"VALUES('" & item_Number & "', " & cost & ", " & quantity_Needed & ", '" & process & "', '" & name_Prefix & "', '" & name & "', '" & type & "', '" & item_Prefix & "')"
					myCmd.ExecuteNonQuery()

					If lastAssembly <> name_Prefix & ":" & name then
						lastAssembly = name_Prefix & ":" & name
						UpdateRTB("    Imported: " & lastAssembly & vbNewLine)
					End If
				Next

				transaction.Commit()
			Catch ex As Exception
				UpdateRTB("       " & ex.Message & vbNewLine)
				If Not transaction Is Nothing Then
					sqlapi.RollBack(transaction, errorMessage:=New List(Of String))
				End If
				QBmajor = True
			End Try
		Catch ex As Exception
			UpdateRTB("       " & ex.Message & vbNewLine)
			_cn.Close()
			QBmajor = True
		End Try
	End Sub
#End Region

#Region "HELPER FUNCTIONS"
	Private Function GetNthIndex(ByVal stringToSearch As String, ByVal charToFind As Char, ByVal number As Integer) As Integer
		Dim count As Integer = 0
		For i As Integer = 0 To stringToSearch.Length - 1
			If stringToSearch(i) = charToFind Then
				count += 1
				If count = number Then
					Return i
				End If
			End If
		Next
		Return -1
	End Function

	Private Sub UpdateRTB(byref text As String)
		RTB_Results.Focus()

		RTB_Results.AppendText(text)
		RTB_Results.Refresh()
	End Sub

	Private Sub TB_ALPHAIndicatorLight_Click() Handles TB_ALPHAIndicatorLight.Click
		If -1 < (RTB_Results.Find("Starting to import SMT BOM")) Then
			RTB_Results.SelectionStart = RTB_Results.Find("Starting to import SMT BOM")
			RTB_Results.ScrollToCaret()
		End If
	End Sub

	Private Sub TB_ALPHAitemsIndicatorLight_Click() Handles TB_ALPHAitemsIndicatorLight.Click
		If CkB_ALPHAitems.Checked = True Then
			If -1 < (RTB_Results.Find("Starting to import SMT Items")) Then
				RTB_Results.SelectionStart = RTB_Results.Find("Starting to import SMT Items")
				RTB_Results.ScrollToCaret()
			End If
		Else
			If -1 < (RTB_Results.Find("Skipping The import of SMT Items")) Then
				RTB_Results.SelectionStart = RTB_Results.Find("Skipping The import of SMT Items")
				RTB_Results.ScrollToCaret()
			End If
		End If
	End Sub

	Private Sub TB_PCADIndicatorLight_Click() Handles TB_PCADIndicatorLight.Click
		If -1 < (RTB_Results.Find("Starting to import PCAD BOM")) Then
			RTB_Results.SelectionStart = RTB_Results.Find("Starting to import PCAD BOM")
			RTB_Results.ScrollToCaret()
		End If
	End Sub

	Private Sub TB_QBIndicatorLight_Click() Handles TB_QBIndicatorLight.Click
		If -1 < (RTB_Results.Find("Starting to import QB BOMs")) Then
			RTB_Results.SelectionStart = RTB_Results.Find("Starting to import QB BOMs")
			RTB_Results.ScrollToCaret()
		End If
	End Sub

	Private Sub TB_QBitemsIndicatorLight_Click() Handles TB_QBitemsIndicatorLight.Click
		If -1 < (RTB_Results.Find("Starting to import QB Items")) Then
			RTB_Results.SelectionStart = RTB_Results.Find("Starting to import QB Items")
			RTB_Results.ScrollToCaret()
		End If
	End Sub
#End Region

End Class