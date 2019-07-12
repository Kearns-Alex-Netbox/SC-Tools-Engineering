'-----------------------------------------------------------------------------------------------------------------------------------------
' Module: AQL_API.vb
'
' Description: SQL helper API for basic SQL functions.
'-----------------------------------------------------------------------------------------------------------------------------------------
Imports System.Data.SqlClient

Public Class SQL_API
	Private username As String = ""
	Private password As String = ""

	''' <summary>
	''' Either sets or returns the string value for username.
	''' </summary>
	''' <value>String value that you want to set username to.</value>
	''' <returns>Returns the string value that username is currently set to.</returns>
	''' <remarks>This is your username.</remarks>
	Public Property _Username() As String
		Get
			Return username.ToString
		End Get
		Set(ByVal value As String)
			username = value
		End Set
	End Property

	''' <summary>
	''' Either sets or returns the string value for password.
	''' </summary>
	''' <value>String value that you want to set password to.</value>
	''' <returns>Returns the string value that password is currently set to.</returns>
	''' <remarks>This is your passwrod.</remarks>
	Public Property _Password() As String
		Get
			Return password.ToString
		End Get
		Set(ByVal value As String)
			password = value
		End Set
	End Property

	''' <summary>
	''' Opens the connection to the database and saves the user who is logged in.
	''' </summary>
	''' <param name="myConn">The connection that you would like to make.</param>
	''' <param name="myCmd">The SQL command that will be used to make the requests.</param>
	''' <returns>True: successful open and return username. False: unsuccessful, see result message for details.</returns>
	''' <remarks>This needs to be called before you make anyother commands.</remarks>
	Public Function OpenDatabase(ByRef myConn As SqlConnection, ByRef myCmd As SqlCommand) As Boolean
		Dim useDatabase As String = ""
		If DATABASE = "Production" Then
			useDatabase = "BOMCompare"
		Else
			useDatabase = "BOMDevel"
		End If
		myConn = New SqlConnection("server=tcp:nas1,1622;Database=" & useDatabase & ";User ID=" & username & ";password= " & password & ";")

		Try
			myConn.Open()
			myCmd = myConn.CreateCommand
			Return True
		Catch ex As Exception
			MsgBox(ex.Message)
			Return False
		End Try
	End Function

	''' <summary>
	''' Closes the database connection that gets passed through.
	''' </summary>
	''' <param name="myConn">The connection that you want to close.</param>
	''' <param name="result">OUTPUT: Error result if somthing does not go right.</param>
	''' <returns>True: Successful close. False: unsuccessful close, see result message for information.</returns>
	''' <remarks>Make sure the connection is already open first before trying to close it.</remarks>
	Public Function CloseDatabase(ByRef myConn As SqlConnection, ByRef result As String) As Boolean
		Try
			If myConn.State <> ConnectionState.Closed Then
				myConn.Close()
			End If
			Return True
		Catch ex As Exception
			result = ex.Message
			Return False
		End Try
	End Function

	''' <summary>
	''' Clears the database of all records.
	''' </summary>
	''' <param name="myCmd">The SQL Command object that we will be using.</param>
	''' <param name="result">The detailed error message if things do not work.</param>
	''' <returns>True: Everything was deleted. False: Something went wrong, see results for more details.</returns>
	''' <remarks></remarks>
	Public Function ClearDatabase(ByRef myCmd As SqlCommand, ByRef result As String, ByRef skipQB As Boolean, ByRef clearALPHA As Boolean) As Boolean
		Try
			myCmd.CommandText = "DELETE FROM " & TABLE_PCADBOM
			myCmd.ExecuteNonQuery()
			myCmd.CommandText = "DELETE FROM " & TABLE_ALPHABOM
			myCmd.ExecuteNonQuery()
			If skipQB = False Then
				myCmd.CommandText = "DELETE FROM " & TABLE_QBBOM
				myCmd.ExecuteNonQuery()
				myCmd.CommandText = "DELETE FROM " & TABLE_QB_ITEMS
				myCmd.ExecuteNonQuery()
			End If
			If clearALPHA = True Then
				myCmd.CommandText = "DELETE FROM " & TABLE_ALPHA_ITEMS
				myCmd.ExecuteNonQuery()
				myCmd.CommandText = "DELETE FROM " & TABLE_MAGAZINE_DATA
				myCmd.ExecuteNonQuery()
				myCmd.CommandText = "DELETE FROM " & TABLE_ALPHA_PACKAGE
				myCmd.ExecuteNonQuery()
			End If

			result = "Database was successfully cleared of all existing data."
			Return True
		Catch ex As Exception
			result = ex.Message
			Return False
		End Try
	End Function

	''' <summary>
	''' Roll back the transaction so we do not commit anything into the database.
	''' </summary>
	''' <param name="transaction">The transaction that we want to roll back.</param>
	''' <param name="errorMessage">OUTPUT: If there is an issue trying to roll back the transaction.</param>
	''' <remarks></remarks>
	Public Sub RollBack(ByRef transaction As SqlTransaction, ByRef errorMessage As List(Of String))
		Try
			If Not transaction Is Nothing Then
				'Attempt to roll back the transaction. 
				transaction.Rollback()
			End If
		Catch ex As Exception
			'Handles any errors that may have occurred on the server that would cause the rollback to fail, such as a closed connection.
			errorMessage.Add(ex.Message)
		End Try
	End Sub

	Public Sub SetDirtyBit(ByRef state As Integer)
		Dim myCmd As New SqlCommand("", myConn)
		If state = 0 Then
			'Turn off the dirty bit
			myCmd.CommandText = "UPDATE " & TABLE_UTILITIES & " SET [" & DB_HEADER_VALUE & "] = 0 WHERE [" & DB_HEADER_NAME & "] = 'Dirty Bit'"
			myCmd.ExecuteNonQuery()
		Else
			'Turn on the dirty bit
			myCmd.CommandText = "UPDATE " & TABLE_UTILITIES & " SET [" & DB_HEADER_VALUE & "] = 1 WHERE [" & DB_HEADER_NAME & "] = 'Dirty Bit'"
			myCmd.ExecuteNonQuery()
			myCmd.CommandText = "UPDATE " & TABLE_UTILITIES & " SET [" & DB_HEADER_VALUE & "] = '" & sqlapi._Username & "' WHERE [" & DB_HEADER_NAME & "] = 'User'"
			myCmd.ExecuteNonQuery()
			myCmd.CommandText = "UPDATE " & TABLE_UTILITIES & " SET [" & DB_HEADER_VALUE & "] = GETDATE() WHERE [" & DB_HEADER_NAME & "] = 'Start'"
			myCmd.ExecuteNonQuery()
		End If
	End Sub

	Public Function CheckDirtyBit(ByRef message As String) As Boolean
		Dim myCmd As New SqlCommand("SELECT [" & DB_HEADER_VALUE & "] FROM " & TABLE_UTILITIES & " WHERE [" & DB_HEADER_NAME & "] = 'Dirty Bit'", myConn)
		Dim dirtybit As Integer = myCmd.ExecuteScalar

		If dirtybit = 0 Then
			Return False
		Else
			myCmd.CommandText = "SELECT [" & DB_HEADER_VALUE & "] FROM " & TABLE_UTILITIES & " WHERE [" & DB_HEADER_NAME & "] = 'User'"
			Dim user As String = myCmd.ExecuteScalar
			myCmd.CommandText = "SELECT [" & DB_HEADER_VALUE & "] FROM " & TABLE_UTILITIES & " WHERE [" & DB_HEADER_NAME & "] = 'Start'"
			Dim starttime As String = myCmd.ExecuteScalar

			message = "There is an import going on." & vbNewLine & "Start: " & starttime & vbNewLine & "User: " & user
			Return True
		End If
	End Function

	''' <summary>
	''' Gets the number of records that are associated with the sql table that we are working with.
	''' </summary>
	''' <param name="myCmd">The sql Command that will be used.</param>
	''' <param name="command">SQL command. Should be 'SELECT COUNT(*) FROM ______'</param>
	''' <param name="numberOfRecords">OUTPUT: The number of records in the table.</param>
	''' <param name="result">OUTPUT: Error message that gives us a hint on what went wrong.</param>
	''' <returns>True: We go our number of records. False: Something went wrong, see error message for information.</returns>
	''' <remarks></remarks>
	Public Function GetNumberOfRecords(ByRef myCmd As SqlCommand, ByRef command As String, ByRef numberOfRecords As Integer, ByRef result As String) As Boolean
		myCmd.CommandText = command
		Dim myReader As SqlDataReader = myCmd.ExecuteReader()

		If myReader.Read() Then
			'Check to see if we are returned a NULL value.
			If myReader.IsDBNull(0) Then
				result = "[GetNumberOfRecords1] Number of records is NULL."
				myReader.Close()
				Return False
			Else
				numberOfRecords = myReader.GetInt32(0)
			End If
		Else
			'If nothing is returned then it does not exist.
			result = "[GetNumberOfRecords2] Number of records does not exist."
			myReader.Close()
			Return False
		End If
		myReader.Close()
		Return True
	End Function

	Public Function FindFile(ByRef myCmd As SqlCommand, ByRef fileName As String, ByRef tableName As String) As Boolean
		'Look for the file name inside the database.
		Dim columnName As String = ""

		If tableName = TABLE_ALPHABOM = True Then
			columnName = DB_HEADER_BOARD_NAME
		ElseIf tableName = TABLE_QBBOM = True Then
			columnName = DB_HEADER_NAME
		Else
			MsgBox("This table is not yet supported with this function.")
			Return False
		End If

		myCmd.CommandText = "IF EXISTS(SELECT [" & columnName & "] FROM " & tableName & " WHERE [" & columnName & "] = '" & fileName & "') SELECT 1 ELSE SELECT 0"
		Dim answer As Integer = myCmd.ExecuteScalar

		If answer = 0 Then
			'The file is not in the database.
			Return False
		Else
			'The file is in the database.
			Return True
		End If
	End Function

#Region "Search Functions with DGVs"

	Public Sub FirstPage(ByRef scrollValue As Integer, ByRef ds As DataSet, ByRef da As SqlDataAdapter, ByRef entriesToShow As Integer)
		scrollValue = 0
		ds.Clear()
		da.Fill(ds, scrollValue, entriesToShow, "Table")
	End Sub

	Public Sub PreviousPage(ByRef scrollValue As Integer, ByRef entriesToShow As Integer, ByRef ds As DataSet, ByRef da As SqlDataAdapter,
							ByRef firstButton As Button, ByRef previousButton As Button)
		scrollValue = scrollValue - entriesToShow
		If scrollValue <= 0 Then
			scrollValue = 0
			firstButton.Enabled = False
			previousButton.Enabled = False
		End If
		ds.Clear()
		da.Fill(ds, scrollValue, entriesToShow, "TABLE")
	End Sub

	Public Sub NextPage(ByRef scrollValue As Integer, ByRef entriesToShow As Integer, ByRef numberOfRecords As Integer, ByRef ds As DataSet,
						ByRef da As SqlDataAdapter, ByRef nextButton As Button, ByRef lastButton As Button)
		scrollValue = scrollValue + entriesToShow
		If scrollValue > numberOfRecords Then
			scrollValue = numberOfRecords - entriesToShow
			If scrollValue < 0 Then
				scrollValue = 0
			End If
			nextButton.Enabled = False
			lastButton.Enabled = False
		End If
		ds.Clear()
		da.Fill(ds, scrollValue, entriesToShow, "TABLE")

		'Check to see if we can keep scrolling or if we are at the end of the table.
		If scrollValue + entriesToShow >= numberOfRecords Then
			nextButton.Enabled = False
			lastButton.Enabled = False
		End If
	End Sub

	Public Sub LastPage(ByRef scrollValue As Integer, ByRef entriesToShow As Integer, ByRef numberOfRecords As Integer, ByRef ds As DataSet, ByRef da As SqlDataAdapter)
		scrollValue = numberOfRecords - entriesToShow
		ds.Clear()
		da.Fill(ds, scrollValue, entriesToShow, "TABLE")
	End Sub

	Public Sub ListAll(ByRef freeze As Integer, ByRef searchCommand As String, ByRef searchCountCommand As String, ByRef myCmd As SqlCommand, ByRef countCommand As String,
					   ByRef numberOfRecords As Integer, ByRef resultsLabel As Label, ByRef command As String, ByRef ds As DataSet, ByRef da As SqlDataAdapter,
					   ByRef dgv As DataGridView, ByRef scrollValue As Integer, ByRef entriesToShow As Integer, ByRef nextButton As Button, ByRef lastButton As Button,
					   ByRef firstButton As Button, ByRef previousButton As Button)
		Dim message As String = ""
		If sqlapi.CheckDirtyBit(message) = True Then
			MsgBox(message)
			Return
		End If

		Dim result As String = ""
		Try
			searchCommand = ""
			searchCountCommand = ""
			sqlapi.GetNumberOfRecords(myCmd, countCommand, numberOfRecords, result)
			resultsLabel.Text = "Number of results: " & numberOfRecords

			myCmd.CommandText = command
			da = New SqlDataAdapter(myCmd)
			ds = New DataSet()

			RetriveData(freeze, da, ds, dgv, scrollValue, entriesToShow, numberOfRecords, nextButton, lastButton, firstButton, previousButton)
		Catch ex As Exception
			MsgBox(ex.Message)
		End Try
	End Sub

	Public Sub Search(ByRef freeze As Integer, ByRef searchTB As TextBox, ByRef searchOperand As ComboBox, ByRef searchCommand As String, ByRef command As String, ByRef searchCountCommand As String, ByRef countCommand As String,
					  ByRef search2TB As TextBox, ByRef searchOperand2 As ComboBox, ByRef searchComboBox As ComboBox, ByRef search2ComboBox As ComboBox, ByRef myCmd As SqlCommand, ByRef ds As DataSet,
					  ByRef da As SqlDataAdapter, ByRef dgv As DataGridView, ByRef numberOfRecords As Integer, ByRef resultsLabel As Label, ByRef scrollValue As Integer,
					  ByRef entriesToShow As Integer, ByRef nextButton As Button, ByRef lastButton As Button, ByRef firstButton As Button, ByRef previousButton As Button,
					  ByRef orderBy As String)
		Dim message As String = ""
		If sqlapi.CheckDirtyBit(message) = True Then
			MsgBox(message)
			Return
		End If

		Dim result As String = ""
		If searchTB.Text.Length = 0 Then
			MsgBox("Please put in a search inside the first box.")
			Return
		End If
		Try
			searchCommand = command
			searchCountCommand = countCommand

			If command.Contains("WHERE") Then
				searchCommand = searchCommand & " AND ("
				Else
				searchCommand = searchCommand & " WHERE ("
				End If

			Select Case searchOperand.Text
				Case "LIKE"
					searchCommand = searchCommand & "[" & searchComboBox.Text & "] LIKE '%" & searchTB.Text & "%'"
				Case "="
					searchCommand = searchCommand & "[" & searchComboBox.Text & "] = '" & searchTB.Text & "'"
				Case "<="
					searchCommand = searchCommand & "[" & searchComboBox.Text & "] <= " & searchTB.Text
				Case ">="
					searchCommand = searchCommand & "[" & searchComboBox.Text & "] >= " & searchTB.Text
			End Select

				If search2TB.Text.Length <> 0 Then
				Select Case searchOperand2.Text
					Case "LIKE"
						searchCommand = searchCommand & " AND [" & search2ComboBox.Text & "] LIKE '%" & search2TB.Text & "%'"
					Case "="
						searchCommand = searchCommand & " AND [" & search2ComboBox.Text & "] = '" & search2TB.Text & "'"
					Case "<="
						searchCommand = searchCommand & " AND [" & search2ComboBox.Text & "] <= " & search2TB.Text
					Case ">="
						searchCommand = searchCommand & " AND [" & search2ComboBox.Text & "] >= " & search2TB.Text
				End Select
			Else
				If searchComboBox.Text = DB_HEADER_ITEM_NUMBER Then
					searchCommand = searchCommand & ") OR ([" & DB_HEADER_MPN & "] LIKE '%" & searchTB.Text & "%' OR [" & DB_HEADER_MPN2 & "] LIKE '%" & searchTB.Text & "%' OR [" & DB_HEADER_MPN3 & "] LIKE '%" & searchTB.Text & "%'"
				End If
			End If



			searchCommand = searchCommand & ") ORDER BY " & orderBy & " ASC"

			If countCommand.Contains("WHERE") Then
				searchCountCommand = searchCountCommand & " AND ("
			Else
				searchCountCommand = searchCountCommand & " WHERE ("
				End If

			Select Case searchOperand.Text
				Case "LIKE"
					searchCountCommand = searchCountCommand & "[" & searchComboBox.Text & "] LIKE '%" & searchTB.Text & "%'"
				Case "="
					searchCountCommand = searchCountCommand & "[" & searchComboBox.Text & "] = '" & searchTB.Text & "'"
				Case "<="
					searchCountCommand = searchCountCommand & "[" & searchComboBox.Text & "] <= " & searchTB.Text
				Case ">="
					searchCountCommand = searchCountCommand & "[" & searchComboBox.Text & "] >= " & searchTB.Text
			End Select

				If search2TB.Text.Length <> 0 Then
				Select Case searchOperand2.Text
					Case "LIKE"
					searchCountCommand = searchCountCommand & " AND [" & search2ComboBox.Text & "] LIKE '%" & search2TB.Text & "%'"
					Case "="
						searchCountCommand = searchCountCommand & " AND [" & search2ComboBox.Text & "] = '" & search2TB.Text & "'"
					Case "<="
						searchCountCommand = searchCountCommand & " AND [" & search2ComboBox.Text & "] <= " & search2TB.Text
					Case ">="
						searchCountCommand = searchCountCommand & " AND [" & search2ComboBox.Text & "] >= " & search2TB.Text
				End Select
			Else
				If searchComboBox.Text = DB_HEADER_ITEM_NUMBER Then
					searchCountCommand = searchCountCommand & ") OR ([" & DB_HEADER_MPN & "] LIKE '%" & searchTB.Text & "%' OR [" & DB_HEADER_MPN2 & "] LIKE '%" & searchTB.Text & "%' OR [" & DB_HEADER_MPN3 & "] LIKE '%" & searchTB.Text & "%'"
				End If
			End If

			searchCountCommand = searchCountCommand & ")"

			sqlapi.GetNumberOfRecords(myCmd, searchCountCommand, numberOfRecords, result)

			resultsLabel.Text = "Number of results: " & numberOfRecords

			myCmd.CommandText = searchCommand
			da = New SqlDataAdapter(myCmd)
			ds = New DataSet()

			RetriveData(freeze, da, ds, dgv, scrollValue, entriesToShow, numberOfRecords, nextButton, lastButton, firstButton, previousButton)
		Catch ex As Exception
			MsgBox(ex.Message)
		End Try
	End Sub

	Public Sub Sort(ByRef freeze As Integer, ByRef searchCommand As String, ByRef command As String, ByRef sortComboBox As ComboBox, ByRef sortASCRadioButton As RadioButton, ByRef myCmd As SqlCommand,
					ByRef ds As DataSet, ByRef da As SqlDataAdapter, ByRef dgv As DataGridView, ByRef scrollValue As Integer, ByRef entriesToShow As Integer,
					ByRef numberOfRecords As Integer, ByRef nextButton As Button, ByRef lastButton As Button, ByRef firstButton As Button, ByRef previousButton As Button)
		Dim message As String = ""
		If sqlapi.CheckDirtyBit(message) = True Then
			MsgBox(message)
			Return
		End If

		Try
			Dim newCommand As String = ""

			If String.Compare(searchCommand, "") = 0 Then
				newCommand = command
			Else
				newCommand = searchCommand
			End If

			If newCommand.Contains("ORDER BY") Then
				newCommand = newCommand.Substring(0, newCommand.IndexOf("ORDER BY ") + 9) & "[" & sortComboBox.Text & "]"   '9 is how long 'ORDER BY ' is
			Else
				newCommand = newCommand & " ORDER BY [" & sortComboBox.Text & "]"
			End If

			If sortASCRadioButton.Checked Then
				newCommand = newCommand & " ASC"
			Else
				newCommand = newCommand & " DESC"
			End If

			myCmd.CommandText = newCommand
			da = New SqlDataAdapter(myCmd)
			ds = New DataSet()

			RetriveData(freeze, da, ds, dgv, scrollValue, entriesToShow, numberOfRecords, nextButton, lastButton, firstButton, previousButton)
		Catch ex As Exception
			MsgBox(ex.Message)
		End Try
	End Sub

	Public Sub RetriveData(ByRef freeze As Integer, ByRef da As SqlDataAdapter, ByRef ds As DataSet, ByRef dgv As DataGridView, ByRef scrollValue As Integer, ByRef entriesToShow As Integer,
						   ByRef numberOfRecords As Integer, ByRef nextButton As Button, ByRef lastButton As Button, ByRef firstButton As Button, ByRef previousButton As Button)
		scrollValue = 0
		da.Fill(ds, scrollValue, entriesToShow, "TABLE")

		If numberOfRecords <= entriesToShow Then
			nextButton.Enabled = False
			lastButton.Enabled = False
		Else
			nextButton.Enabled = True
			lastButton.Enabled = True
		End If
		previousButton.Enabled = False
		firstButton.Enabled = False

		dgv.DataSource = Nothing
		dgv.DataSource = ds.Tables("TABLE")
		dgv.AutoResizeColumns(DataGridViewAutoSizeColumnMode.AllCells)
		dgv.Columns(freeze).Frozen = True
	End Sub

#End Region

End Class