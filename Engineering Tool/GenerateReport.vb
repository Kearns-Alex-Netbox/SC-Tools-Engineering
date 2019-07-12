'-----------------------------------------------------------------------------------------------------------------------------------------
' Module: GenerateReport.vb
'
' Description: Generates each report through Excel with special formatting.
'
'-----------------------------------------------------------------------------------------------------------------------------------------
Imports Microsoft.Office.Interop

Public Class GenerateReport

	Public Sub GenerateImportReport(ByRef list As RichTextBox)
		Try
			Dim xlApp As New Excel.Application
			Dim xlWorkBook As Excel.Workbook
			Dim xlWorkSheet As Excel.Worksheet
			Dim misValue As Object = Reflection.Missing.Value
			Dim INDEX_row As Integer = 1
			Dim INDEX_column As Integer = 1

			xlWorkBook = xlApp.Workbooks.Add(misValue)
			xlWorkSheet = xlWorkBook.Sheets("sheet1")

			xlWorkSheet.PageSetup.CenterHeader = "Import Output Report   " & Date.Now

			For Each line In list.Lines
				xlWorkSheet.Cells(INDEX_row, 1) = line
				INDEX_row += 1
			Next

			xlWorkSheet.Range("A1:X1").EntireColumn.AutoFit()
			xlWorkSheet.Range("A1:X1").EntireColumn.NumberFormat = "0"
			xlWorkSheet.Range("A1:X1").EntireColumn.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft

			xlApp.DisplayAlerts = False
			xlApp.Visible = True

			releaseObject(xlWorkSheet)
			releaseObject(xlWorkBook)
			releaseObject(xlApp)
		Catch ex As Exception
			MsgBox("File was not written: " & ex.Message)
		End Try
	End Sub

	Public Sub GenerateRevisionReport(ByRef source As String, ByRef board1 As String, ByRef board2 As String,
									  ByRef board1Quantity As DataTable, ByRef board2Quantity As DataTable)
		Try
			Dim xlApp As New Excel.Application
			Dim xlWorkBook As Excel.Workbook
			Dim xlWorkSheet As Excel.Worksheet
			Dim misValue As Object = Reflection.Missing.Value
			xlWorkBook = xlApp.Workbooks.Add(misValue)

			'----- SHEET 1 -----'

			xlWorkSheet = xlWorkBook.Sheets("sheet1")
			xlWorkSheet.Name = "Quantities"
			xlWorkSheet.PageSetup.CenterHeader = source & " Revision Quantity Report for: " & board1 & " - " & board2 & "   " & Date.Now.Date

			'ROW 1
			Dim INDEX_row As Integer = 1
			Dim INDEX_column As Integer = 1

			'ROW 2
			INDEX_row += 1

			For Each header In board1Quantity.Columns
				xlWorkSheet.Cells(INDEX_row, INDEX_column) = header.columnName
				INDEX_column += 1
			Next

			xlWorkSheet.Cells(1, 1) = board1 & " Against " & board2
			xlWorkSheet.Range(xlWorkSheet.Cells(1, 1), xlWorkSheet.Cells(1, INDEX_column - 1)).MergeCells = True

			INDEX_column += 1
			Dim nextColumn2 = INDEX_column

			For Each header In board2Quantity.Columns
				xlWorkSheet.Cells(INDEX_row, INDEX_column) = header.columnName
				INDEX_column += 1
			Next

			xlWorkSheet.Cells(1, nextColumn2) = board2 & " Against " & board1
			xlWorkSheet.Range(xlWorkSheet.Cells(1, nextColumn2), xlWorkSheet.Cells(1, INDEX_column - 1)).MergeCells = True

			'ROW 3
			INDEX_row += 1
			INDEX_column = 1

			For row = 0 To board1Quantity.Rows.Count - 1
				For column = 0 To board1Quantity.Columns.Count - 1
					xlWorkSheet.Cells(INDEX_row, INDEX_column) = board1Quantity(row)(column)
					INDEX_column += 1
				Next
				INDEX_column = 1
				INDEX_row += 1
			Next

			INDEX_row = 3
			INDEX_column = nextColumn2

			For row = 0 To board2Quantity.Rows.Count - 1
				For column = 0 To board2Quantity.Columns.Count - 1
					xlWorkSheet.Cells(INDEX_row, INDEX_column) = board2Quantity(row)(column)
					INDEX_column += 1
				Next
				INDEX_column = nextColumn2
				INDEX_row += 1
			Next

			xlWorkSheet.Range("A1:X1").EntireColumn.AutoFit()
			xlWorkSheet.Range("A1:X1").EntireColumn.NumberFormat = "0"
			xlWorkSheet.Range("A1:A2").EntireRow.Font.Bold = True
			xlWorkSheet.Range("A1:X1").EntireColumn.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft

			xlApp.DisplayAlerts = False
			xlApp.Visible = True

			releaseObject(xlWorkSheet)
			releaseObject(xlWorkBook)
			releaseObject(xlApp)
		Catch ex As Exception
			MsgBox("File was not written: " & ex.Message)
		End Try
	End Sub

	Public Sub GenerateBoardOptionReport(ByRef table As DataTable)
		Try
			Dim xlApp As New Excel.Application
			Dim xlWorkBook As Excel.Workbook
			Dim xlWorkSheet As Excel.Worksheet
			Dim misValue As Object = Reflection.Missing.Value
			Dim INDEX_row As Integer = 1
			Dim INDEX_column As Integer = 1

			xlWorkBook = xlApp.Workbooks.Add(misValue)
			xlWorkSheet = xlWorkBook.Sheets("sheet1")

			xlWorkSheet.PageSetup.CenterHeader = "Board Option Report   " & Date.Now

			For Each header In table.Columns
				xlWorkSheet.Cells(INDEX_row, INDEX_column) = header.columnName
				INDEX_column += 1
			Next

			INDEX_column = 1
			INDEX_row += 1
			For row = 0 To table.Rows.Count - 1
				For column = 0 To table.Columns.Count - 1
					xlWorkSheet.Cells(INDEX_row, INDEX_column) = table(row)(column)
					INDEX_column += 1
				Next
				INDEX_column = 1
				INDEX_row += 1
			Next

			xlWorkSheet.Range("A1:C1").EntireColumn.AutoFit()
			xlWorkSheet.Range("A1:X1").EntireColumn.NumberFormat = "0"
			xlWorkSheet.Range("A1").EntireRow.Font.Bold = True
			xlWorkSheet.Range("A1:X1").EntireColumn.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft

			xlApp.DisplayAlerts = False
			xlApp.Visible = True

			releaseObject(xlWorkSheet)
			releaseObject(xlWorkBook)
			releaseObject(xlApp)
		Catch ex As Exception
			MsgBox("File was not written: " & ex.Message)
		End Try
	End Sub

	Public Sub GeneratePartUsageReport(ByRef table As DataTable)
		Try
			Dim xlApp As New Excel.Application
			Dim xlWorkBook As Excel.Workbook
			Dim xlWorkSheet As Excel.Worksheet
			Dim misValue As Object = Reflection.Missing.Value
			Dim INDEX_row As Integer = 1
			Dim INDEX_column As Integer = 1

			Dim costIndex As Integer = 0

			xlWorkBook = xlApp.Workbooks.Add(misValue)
			xlWorkSheet = xlWorkBook.Sheets("sheet1")

			xlWorkSheet.PageSetup.CenterHeader = "Part Usage Report   " & Date.Now

			For Each header In table.Columns
				xlWorkSheet.Cells(INDEX_row, INDEX_column) = header.columnName
				If header.ColumnName = DB_HEADER_COST Then
					costIndex = INDEX_column
				End If
				INDEX_column += 1
			Next

			INDEX_column = 1
			INDEX_row += 1
			For row = 0 To table.Rows.Count - 1
				For column = 0 To table.Columns.Count - 1
					xlWorkSheet.Cells(INDEX_row, INDEX_column) = table(row)(column)
					INDEX_column += 1
				Next
				INDEX_column = 1
				INDEX_row += 1
			Next

			xlWorkSheet.Range("A1:AY100").EntireColumn.AutoFit()
			xlWorkSheet.Range("A1:AY100").EntireColumn.NumberFormat = "0"
			xlWorkSheet.Range("A1").EntireRow.Font.Bold = True

			'Cost
			If costIndex <> 0 Then
				xlWorkSheet.Cells(1, costIndex).EntireColumn.NumberFormat = "_($* #,##0.00#####_);_($* (#,##0.00#####);_($* ""-""??_);_(@_)"
			End If

			Dim range As Excel.Range
			range = xlWorkSheet.UsedRange
			Dim borders As Excel.Borders = range.Borders
			borders.LineStyle = Excel.XlLineStyle.xlContinuous

			xlWorkSheet.Range("A1:AY1").EntireColumn.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft

			xlApp.DisplayAlerts = False
			xlApp.Visible = True

			releaseObject(xlWorkSheet)
			releaseObject(xlWorkBook)
			releaseObject(xlApp)
		Catch ex As Exception
			MsgBox("File was not written: " & ex.Message)
		End Try
	End Sub

	Public Sub GenerateQB_itemslistReport(ByRef ds As DataSet)
		Try
			Dim xlApp As New Excel.Application
			Dim xlWorkBook As Excel.Workbook
			Dim xlWorkSheet As Excel.Worksheet
			Dim misValue As Object = Reflection.Missing.Value
			Dim INDEX_row As Integer = 1
			Dim INDEX_column As Integer = 1

			xlWorkBook = xlApp.Workbooks.Add(misValue)
			xlWorkSheet = xlWorkBook.Sheets("sheet1")

			xlWorkSheet.PageSetup.CenterHeader = "QB Items Report   " & Date.Now

			For Each dc As DataColumn In ds.Tables(0).Columns
				xlWorkSheet.Cells(1, INDEX_column) = dc.ColumnName
				INDEX_column += 1
			Next

			INDEX_row += 1
			'Reset the Column index
			INDEX_column = 1

			For Each dr As DataRow In ds.Tables(0).Rows
				For Each dc As DataColumn In ds.Tables(0).Columns
					xlWorkSheet.Cells(INDEX_row, INDEX_column) = dr(dc).ToString
					INDEX_column += 1
				Next
				INDEX_row += 1
				'Reset the Column index
				INDEX_column = 1
			Next

			xlWorkSheet.Range("A1:X1").EntireColumn.AutoFit()
			xlWorkSheet.Range("A1:X1").EntireColumn.NumberFormat = "0"
			xlWorkSheet.Range("A1").EntireRow.Font.Bold = True
			xlWorkSheet.Range("A1:X1").EntireColumn.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft

			xlApp.DisplayAlerts = False
			xlApp.Visible = True

			releaseObject(xlWorkSheet)
			releaseObject(xlWorkBook)
			releaseObject(xlApp)
		Catch ex As Exception
			MsgBox("File was not written: " & ex.Message)
		End Try
	End Sub

	Public Sub GenerateBOMCompareReport(ByRef dsPCAD As DataSet, ByRef dsQB As DataSet, ByRef boardName As String)
		Try
			Dim xlApp As New Excel.Application
			Dim xlWorkBook As Excel.Workbook
			Dim xlWorkSheet As Excel.Worksheet
			Dim misValue As Object = Reflection.Missing.Value
			Dim INDEX_row As Integer = 1
			Dim INDEX_column As Integer = 1

			xlWorkBook = xlApp.Workbooks.Add(misValue)
			xlWorkSheet = xlWorkBook.Sheets("sheet1")

			xlWorkSheet.PageSetup.CenterHeader = "PCAD:  " & boardName & " Report   " & Date.Now

			For Each dc As DataColumn In dsPCAD.Tables(0).Columns
				xlWorkSheet.Cells(1, INDEX_column) = dc.ColumnName
				INDEX_column += 1
			Next

			INDEX_row += 1
			'Reset the Column index
			INDEX_column = 1

			For Each dr As DataRow In dsPCAD.Tables(0).Rows
				For Each dc As DataColumn In dsPCAD.Tables(0).Columns
					xlWorkSheet.Cells(INDEX_row, INDEX_column) = dr(dc).ToString
					INDEX_column += 1
				Next
				INDEX_row += 1
				'Reset the Column index
				INDEX_column = 1
			Next

			xlWorkSheet.Range("A1:X1").EntireColumn.AutoFit()
			xlWorkSheet.Range("A1:X1").EntireColumn.NumberFormat = "0"
			xlWorkSheet.Range("A1").EntireRow.Font.Bold = True
			xlWorkSheet.Range("A1:X1").EntireColumn.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft

			If Not dsQB Is Nothing Then
				INDEX_row = 1
				INDEX_column = 1
				xlWorkSheet = xlWorkBook.Sheets("sheet2")

				xlWorkSheet.PageSetup.CenterHeader = "QB:  " & boardName & " Report   " & Date.Now

				For Each dc As DataColumn In dsQB.Tables(0).Columns
					xlWorkSheet.Cells(1, INDEX_column) = dc.ColumnName
					INDEX_column += 1
				Next

				INDEX_row += 1
				'Reset the Column index
				INDEX_column = 1

				For Each dr As DataRow In dsQB.Tables(0).Rows
					For Each dc As DataColumn In dsQB.Tables(0).Columns
						xlWorkSheet.Cells(INDEX_row, INDEX_column) = dr(dc).ToString
						INDEX_column += 1
					Next
					INDEX_row += 1
					'Reset the Column index
					INDEX_column = 1
				Next

				xlWorkSheet.Range("A1:X1").EntireColumn.AutoFit()
				xlWorkSheet.Range("A1:X1").EntireColumn.NumberFormat = "0"
				xlWorkSheet.Range("A1").EntireRow.Font.Bold = True
				xlWorkSheet.Range("A1:X1").EntireColumn.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft
			End If

			xlApp.DisplayAlerts = False
			xlApp.Visible = True

			releaseObject(xlWorkSheet)
			releaseObject(xlWorkBook)
			releaseObject(xlApp)
		Catch ex As Exception
			MsgBox("File was not written: " & ex.Message)
		End Try
	End Sub

	Public Function ModifyExcel(ByVal txtFilePath As String, ByVal boardName As String) As Boolean
		' setup the workbook
		Dim xlApp As New Excel.Application
		Dim xlWorkBook As Excel.Workbook
		Dim xlWorkSheet As Excel.Worksheet
		Dim misValue As Object = Reflection.Missing.Value
		Dim INDEX_row As Integer = 1
		Dim INDEX_column As Integer = 1

		Try
			xlWorkBook = xlApp.Workbooks.Open(My.Settings.RevisionFile)

			Dim sheetFound As Boolean = False

			'check to see that we have already added this sheet
			For Each xs In xlWorkBook.Sheets
				If xs.Name = boardName Then
					sheetFound = True
					Exit For
				End If
			Next

			If sheetFound = False Then
				' add the sheet
				xlWorkSheet = xlWorkBook.Sheets.Add(After:=xlWorkBook.Sheets(xlWorkBook.Sheets.Count))
				xlWorkSheet.Name = boardName
				xlWorkSheet.PageSetup.CenterHeader = boardName
			End If

			xlWorkSheet = xlWorkBook.Sheets(boardName)

			'find the lowest row that we have not used yet
			Do
				' We want to have two consecutive rows that are blank. This will allow us to have spacing inbetween revisions
				If String.IsNullOrEmpty(xlWorkSheet.Cells(INDEX_row, 1).Value) And String.IsNullOrEmpty(xlWorkSheet.Cells(INDEX_row, 2).Value) And
				   String.IsNullOrEmpty(xlWorkSheet.Cells(INDEX_row + 1, 1).Value) Then
					Exit Do
				End If

				INDEX_row += 1
			Loop


			' Open the text file
			Using sr As New IO.StreamReader(txtFilePath)
				Dim line As String = ""

				Do
					line = sr.ReadLine()
					If line Is Nothing Then
						Exit Do
					End If

					' if we hit a blank line then just continue
					If line.Length = 0 Then
						Continue Do
					End If

					Select Case line.Substring(0, 1)
						Case "-"
							INDEX_row += 1
							xlWorkSheet.Cells(INDEX_row, 1) = line
							xlWorkSheet.Cells(INDEX_row, 1).Font.Bold = True
							xlWorkSheet.Cells(INDEX_row, 1).Font.size = 14
						Case "*"
							xlWorkSheet.Cells(INDEX_row, 1) = line
						Case Else
							xlWorkSheet.Cells(INDEX_row, 2) = line
					End Select

					INDEX_row += 1
				Loop
			End Using

			xlApp.DisplayAlerts = False
			xlApp.Visible = True
			xlWorkBook.Save()

			releaseObject(xlWorkSheet)
			releaseObject(xlWorkBook)
			releaseObject(xlApp)
		Catch ex As Exception
			MsgBox(ex.Message)
			releaseObject(xlWorkSheet)
			releaseObject(xlWorkBook)
			releaseObject(xlApp)
			Return False
		End Try

		Return True
	End Function

	Private Sub releaseObject(ByVal obj As Object)
		Try
			Runtime.InteropServices.Marshal.ReleaseComObject(obj)
			obj = Nothing
		Catch ex As Exception
			obj = Nothing
		Finally
			GC.Collect()
		End Try
	End Sub

End Class
