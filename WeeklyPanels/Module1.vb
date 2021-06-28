Imports Microsoft.Office.Interop.Excel

Module Module1
    Public Const APPNAME = "WeeklyPanels"
    Public Const ROOTDIR = "O:\"
    Public Const ROOTPANELS = ROOTDIR & "PANELS\"
    Public Const NEWPTOLD = ROOTPANELS & "sep13_2019\new_patients_sep13_2019.xlsx"
    Public Const PANELOLD = ROOTPANELS & "sep13_2019\panel_patients_sep13_2019.xlsx"
    Public Const PANELRECENT = ROOTPANELS & "panel_patients_oct11_2019.xlsx"
    Public Const NEWPTRECENT = ROOTPANELS & "new_patients_oct11_2019.xlsx"
    Public Const PANELREMOVED = ROOTPANELS & "PANEL-Removed"
    Public Const PANELADDED = ROOTPANELS & "PANEL-Added"
    Public Const NEWPTSREMOVED = ROOTPANELS & "NEWPTS-Removed"
    Public Const NEWPTSADDED = ROOTPANELS & "NEWPTS-Added"
    Public Const NAMECHANGED = ROOTPANELS & "NAME_CHANGED"

    Public fileLog As System.IO.StreamWriter  ' set in InitializeProcessingLog()
    Dim excelApplication As Application = Nothing

    '--set in fncVerifyExcel()
    Dim wbPanelOld As Workbook = Nothing
    Dim wbPanelRecent As Workbook = Nothing
    Dim wbNewPtOld As Workbook = Nothing
    Dim wbNewPtRecent As Workbook = Nothing

    Sub WMsg(strText As String)

        System.Console.WriteLine(strText)
        fileLog.WriteLine(strText)  ' fileLog set in InitializeProcessingLog

        'Dim intTimer As Integer = 500
        'System.Threading.Thread.Sleep(intTimer)  ' comment out for speed

    End Sub

    Sub InitializeProcessingLog()
        Dim strFileName As String
        strFileName = CStr(Now)
        strFileName = Replace(strFileName, "/", "-")
        strFileName = Replace(strFileName, ":", "-")

        System.Console.WriteLine("Begin Processing " & CStr(Now))

        strFileName = ROOTDIR & "processing_log\" & APPNAME & " " & strFileName & ".txt"

        fileLog = My.Computer.FileSystem.OpenTextFileWriter(strFileName, True)
        fileLog.WriteLine("Program start " & CStr(Now))
    End Sub

    Sub Main()
        Try
            InitializeProcessingLog()  ' open fileLog global
            excelApplication = New Application()
            If Not fncVerifyExcel() Then
                Throw New System.Exception("Excel workbook(s) cannot be verified")
            End If
            RemovedNewPatients()
            RemovedFromPanels()
            NewlyAddedToPanels()
            NewlyAddedNewPatients()

            WMsg("Remove patients whose names changed")
            DeletePatientsWhoseNamesChanged()

        Catch ex As Exception
            WMsg("Main() ERROR " & ex.Message)
        Finally
            WMsg("CLOSING EXCEL...")
            CloseWorkbooks()
            CloseExcel()
            WMsg("PROGRAM END " & CStr(Now))
            fileLog.Close() ' close global
        End Try
    End Sub

    Function fncVerifyExcel() As Boolean
        fncVerifyExcel = False
        Try

            excelApplication.ScreenUpdating = False
            wbNewPtOld = excelApplication.Workbooks.Open(NEWPTOLD)
            wbPanelOld = excelApplication.Workbooks.Open(PANELOLD)
            wbPanelRecent = excelApplication.Workbooks.Open(PANELRECENT)
            wbNewPtRecent = excelApplication.Workbooks.Open(NEWPTRECENT)

            ' Column H must be computed to Column A + Column G before processing can occur
            If wbNewPtOld.Sheets(1).Range("H2").Value = "" Or
                    wbPanelOld.Sheets(1).Range("H2").Value = "" Or
                    wbPanelRecent.Sheets(1).Range("H2").Value = "" Or
                    wbNewPtRecent.Sheets(1).Range("H2").Value = "" Then
                Throw New System.Exception("Missing column H computed value")
            End If

            fncVerifyExcel = True
        Catch ex As Exception
            WMsg("fncVerifyExcel ERROR " & ex.Message)
        End Try
    End Function

    Sub RemovedFromPanels()
        WMsg("RemovedFromPanels()")
        FindMatches(PANELREMOVED, wbPanelOld, wbPanelRecent)
    End Sub
    Sub NewlyAddedToPanels()
        WMsg("NewlyAddedToPanels()")
        FindMatches(PANELADDED, wbPanelRecent, wbPanelOld)
    End Sub
    Sub RemovedNewPatients()
        WMsg("RemovedNewPatients()")
        FindMatches(NEWPTSREMOVED, wbNewPtOld, wbNewPtRecent)
    End Sub
    Sub NewlyAddedNewPatients()
        WMsg("NewlyAddedNewPatients()")
        FindMatches(NEWPTSADDED, wbNewPtRecent, wbNewPtOld)
    End Sub

    Sub FindMatches(strName As String, ByRef wbColA As Workbook, ByRef wbColC As Workbook)
        '--Make a brand new workbook, and then using the wbColA and wbColC workbooks
        '--Compare the values in Column H from each workbook.
        '--The new workbook stores this comparison and all the names NOT THE SAME
        '--will be blank in Column B of this new workbook
        Dim wbNew As Workbook = excelApplication.Workbooks.Add()
        Try
            Dim wsSourceA, wsSourceC, wsDest As Worksheet

            wsDest = wbNew.Sheets(1)

            wsSourceA = wbColA.Sheets(1)
            CopySheetColumn(wsSourceA, "H:H", wsDest, "A1")

            wsSourceC = wbColC.Sheets(1)
            CopySheetColumn(wsSourceC, "H:H", wsDest, "C1")

            CompareColumnsAC(wsDest)

            WMsg("....set up new Worksheet")
            wsSourceA.Copy(After:=wbNew.Sheets(1))
            CopyColBToColI(wbNew)
            FormatList(wbNew.Sheets(2))

            wsDest.Range("O1").Copy()  ' empty the clipboard
        Catch ex As Exception
            WMsg("FindMatches ERROR " & ex.Message)
        Finally
            WMsg("....save newly created Workbook as " & strName)
            wbNew.SaveAs(strName)
            wbNew.Close(False)
        End Try
    End Sub

    Sub CopySheetColumn(ByRef wsSource As Worksheet, strSourceColumn As String,
                        ByRef wsDest As Worksheet, strDestRng As String)
        Try
            wsSource.Range(strSourceColumn).Copy()
            wsDest.Range(strDestRng).PasteSpecial(XlPasteType.xlPasteValues)
        Catch ex As Exception
            WMsg("CopySheetColumn ERROR " & ex.Message)
        End Try
    End Sub

    Sub CompareColumnsAC(ByRef wsDest As Worksheet)

        Try
            '--Excel Interop returns a 1-based array when getting the Range Value (1,1)
            '--All other arrays are 0-based by default

            Dim arrCompareA, arrCompareC As Array  ' these will be 1-based arrays (1,1)

            Dim arrC(0) As String 'Need the 0-based array to use IndexOf

            Dim arrInsertB(,) As String  'Used to write back all values to Excel column B; needs 2 dimensions

            '--all values from ColumnA of wsDest sheet go into arrCompareA 1-based
            arrCompareA = wsDest.Range("A1:A" & fncGetLastRow(wsDest, "A")).Value
            ReDim arrInsertB(UBound(arrCompareA), 0)  ' set the size of this 2-dimension array

            '--all values from ColumnC of wsDest sheet go into arrC 0-based
            arrCompareC = wsDest.Range("C1:C" & fncGetLastRow(wsDest, "C")).Value
            ReDim arrC(UBound(arrCompareC) - 1)  ' set the size of this 0-based 1-dimension array

            '--copy values from the 1-based arrCompareC into the 0-based array arrC
            For x = 1 To UBound(arrCompareC)
                arrC(x - 1) = arrCompareC(x, 1)
            Next

            '--If value in Column A exists in Column C it will be copied into Column B
            For x = 1 To UBound(arrCompareA)
                If Array.IndexOf(arrC, arrCompareA(x, 1)) > -1 Then
                    arrInsertB(x - 1, 0) = arrCompareA(x, 1)
                End If
            Next

            '--Copy all matching values to column B and increase column widths
            wsDest.Range("B1:B" & UBound(arrInsertB)).Value = arrInsertB
            wsDest.Range("A:C").ColumnWidth = 48

        Catch ex As Exception
            WMsg("CompareColumnsAC ERROR " & ex.Message)
        End Try
    End Sub

    Function fncGetLastRow(ByRef wsSheet As Worksheet, strColumn As String) As Long
        With wsSheet
            fncGetLastRow = .Range(strColumn & .Rows.Count).End(XlDirection.xlUp).Row
        End With
    End Function

    Sub CopyColBToColI(ByRef wbNew As Workbook)

        Try
            Dim sourceColumn As Range, targetColumn As Range
            sourceColumn = wbNew.Sheets(1).Columns("B")
            targetColumn = wbNew.Sheets(2).Columns("I")

            sourceColumn.Copy(Destination:=targetColumn)

        Catch ex As Exception
            WMsg("CopyColBToColI ERROR " & ex.Message)
        End Try
    End Sub

    Sub FormatList(ByRef wsSheet As Worksheet)
        '--remove all patient names that stayed the same
        Try
            '--Sort by cols I, G, A
            SortByIGA(wsSheet)

            '--Delete all rows that have a value in col I
            DeleteRowsWithValues(wsSheet, "I")
            wsSheet.Range("A1").Select()
        Catch ex As Exception
            WMsg("FormatList ERROR " & ex.Message)
        End Try
    End Sub

    Sub SortByIGA(ByRef wsSheet As Worksheet)
        Try
            Dim myRange As Range

            myRange = wsSheet.Range("A1", "I" & CStr(fncGetLastRow(wsSheet, "I")))

            myRange.Sort(Key1:=myRange.Range("I1"),
                         Key2:=myRange.Range("G1"),
                         Key3:=myRange.Range("A1"),
                                Order1:=XlSortOrder.xlAscending,
                                Orientation:=XlSortOrientation.xlSortColumns)

        Catch ex As Exception
            WMsg("SortByIGA ERROR " & ex.Message)
        End Try
    End Sub

    Sub SortByColG(ByRef wsSheet As Worksheet)
        Try
            Dim myRange As Range

            myRange = wsSheet.Range("A1", "G" & CStr(fncGetLastRow(wsSheet, "A")))

            myRange.Sort(Key1:=myRange.Range("G1"),
                         Key2:=myRange.Range("A1"),
                                Order1:=XlSortOrder.xlAscending,
                                Orientation:=XlSortOrientation.xlSortColumns)

        Catch ex As Exception
            WMsg("SortByColG ERROR " & ex.Message)
        End Try
    End Sub

    Sub SortByColC(ByRef wsSheet As Worksheet)
        Try
            Dim myRange As Range

            myRange = wsSheet.Range("A1", "E" & CStr(fncGetLastRow(wsSheet, "C")))

            myRange.Sort(Key1:=myRange.Range("C1"),
                                Order1:=XlSortOrder.xlAscending,
                                Orientation:=XlSortOrientation.xlSortColumns)

        Catch ex As Exception
            WMsg("SortByColC ERROR " & ex.Message)
        End Try
    End Sub

    Sub SortByColDA(ByRef wsSheet As Worksheet)
        Try
            Dim myRange As Range

            myRange = wsSheet.Range("A1", "E" & CStr(fncGetLastRow(wsSheet, "E")))

            myRange.Sort(Key1:=myRange.Range("D1"),
                         Key2:=myRange.Range("A1"),
                                Order1:=XlSortOrder.xlAscending,
                                Orientation:=XlSortOrientation.xlSortColumns)

        Catch ex As Exception
            WMsg("SortByColDA ERROR " & ex.Message)
        End Try
    End Sub

    Sub DeleteRowsWithValues(ByRef wsSheet As Worksheet, strColumn As String)
        Try
            Dim strL As String = CStr(fncGetLastRow(wsSheet, strColumn))
            wsSheet.Range(strColumn & "1:" & strColumn &
                          strL).EntireRow.Delete(Shift:=XlDirection.xlUp)
        Catch ex As Exception
            WMsg("DeleteRowsWithValues ERROR " & ex.Message)
        End Try
    End Sub

    Sub DeletePatientsWhoseNamesChanged()
        '--Go into the newly created Panel added/removed workbooks
        '--And delete the names of those patients who just had a name change
        '--i.e. their PHN+PHYSICIAN are still the same
        Try
            Dim wbNew As Workbook = excelApplication.Workbooks.Add
            Dim wbAdded As Workbook = excelApplication.Workbooks.Open(PANELADDED)
            Dim wbRemoved As Workbook = excelApplication.Workbooks.Open(PANELREMOVED)

            '--remove columns and format sizes
            CleanupSheet(wbAdded.Sheets(2))
            CleanupSheet(wbRemoved.Sheets(2))

            '--calculate a column of values for PHN & PHYSICIAN
            CalcPHNPhysVals(wbAdded.Sheets(2))
            CalcPHNPhysVals(wbRemoved.Sheets(2))

            '--see which are same in both Added and Removed 
            FindNameChanges(wbNew, wbAdded, wbRemoved)

            '--move those patients who are name changes into the wbNew Workbook sheet
            RecordNameChanges(wbNew, wbAdded, wbRemoved)

            '--delete names changed in both workbooks wbAdded and wbRemoved
            DeleteNamesChanged(wbAdded.Sheets(2), wbRemoved.Sheets(2))

            wbAdded.Save()
            wbRemoved.Save()
            wbNew.SaveAs(NAMECHANGED)
            wbNew.Close(False)

        Catch ex As Exception
            WMsg("DeletePatientsWhoseNamesChanged ERROR " & ex.Message)
        End Try
    End Sub

    Sub CleanupSheet(ByRef wsSheet As Worksheet)
        Try

            '--Delete column "H"
            wsSheet.Range("H1").EntireColumn.Delete()

            '--Move AGE Column "D" to Column "H"
            wsSheet.Range("D1").EntireColumn.Copy(wsSheet.Range("H1").EntireColumn)

            '--Remove columns "B" to "D"
            wsSheet.Range("B1:D1").EntireColumn.Delete()

            '--Resize
            wsSheet.Range("A1").ColumnWidth = 32
            wsSheet.Range("B1").ColumnWidth = 10.75
            wsSheet.Range("C1").ColumnWidth = 9.5
            wsSheet.Range("D1").ColumnWidth = 18
            wsSheet.Range("E1").Value = "Age"
            wsSheet.Range("E1").ColumnWidth = 4
            wsSheet.Range("F1").ColumnWidth = 44
            wsSheet.Range("G1").ColumnWidth = 44


        Catch ex As Exception
            WMsg("CleanupSheet ERROR " & ex.Message)
        End Try

    End Sub

    Sub CalcPHNPhysVals(ByRef wsSheet As Worksheet)
        Try
            '--Formula for Column F should be Col C & " " & Col D
            For x = 2 To fncGetLastRow(wsSheet, "D")
                wsSheet.Range("F" & CStr(x)).Formula =
                    "=CONCATENATE(C" & CStr(x) & ",D" & CStr(x) & ")"
            Next
        Catch ex As Exception
            WMsg("CalcPHNPhysVals ERROR " & ex.Message)
        End Try
    End Sub

    Sub FindNameChanges(ByRef wbNew As Workbook, ByRef wbAdded As Workbook, ByRef wbRemoved As Workbook)
        '--Look for values in Column F of each workbook that are the same, remove those rows and
        '--record the names being removed in wbNew

        Try
            wbNew.Sheets.Add()
            CopySheetColumn(wbAdded.Sheets(2), "F:F", wbNew.Sheets(1), "A1")
            CopySheetColumn(wbRemoved.Sheets(2), "F:F", wbNew.Sheets(1), "C1")
            CompareColumnsAC(wbNew.Sheets(1))

            CopySheetColumn(wbRemoved.Sheets(2), "F:F", wbNew.Sheets(2), "A1")
            CopySheetColumn(wbAdded.Sheets(2), "F:F", wbNew.Sheets(2), "C1")
            CompareColumnsAC(wbNew.Sheets(2))

            CopyColBToColG(wbNew.Sheets(1), wbAdded.Sheets(2))
            CopyColBToColG(wbNew.Sheets(2), wbRemoved.Sheets(2))

        Catch ex As Exception
            WMsg("FindNameChanges ERROR " & ex.Message)
        End Try
    End Sub

    Sub CopyColBToColG(ByRef wsSource As Worksheet, ByRef wsDest As Worksheet)

        Try
            Dim sourceColumn As Range, targetColumn As Range
            sourceColumn = wsSource.Columns("B")
            targetColumn = wsDest.Columns("G")

            sourceColumn.Copy(Destination:=targetColumn)

        Catch ex As Exception
            WMsg("CopyColBToColG ERROR " & ex.Message)
        End Try
    End Sub

    Sub RecordNameChanges(ByRef wbNew As Workbook, ByRef wbAdded As Workbook, ByRef wbRemoved As Workbook)

        Try
            SortByColG(wbRemoved.Sheets(2))
            Dim wsRecord As Worksheet = wbNew.Sheets.Add()
            For x = 1 To fncGetLastRow(wbRemoved.Sheets(2), "A")
                If Len(wbRemoved.Sheets(2).Range("G" & CStr(x)).Value) > 0 Then
                    wsRecord.Range("A" & CStr(x)).Value =
                    wbRemoved.Sheets(2).Range("A" & CStr(x)).Value
                End If
            Next

            SortByColG(wbAdded.Sheets(2))
            For x = 1 To fncGetLastRow(wbAdded.Sheets(2), "A")
                If Len(wbAdded.Sheets(2).Range("G" & CStr(x)).Value) > 0 Then
                    wsRecord.Range("C" & CStr(x)).Value =
                    wbAdded.Sheets(2).Range("A" & CStr(x)).Value
                    wsRecord.Range("D" & CStr(x)).Value =
                    wbAdded.Sheets(2).Range("D" & CStr(x)).Value
                    wsRecord.Range("E" & CStr(x)).Value =
                    wbAdded.Sheets(2).Range("C" & CStr(x)).Value

                End If
            Next

            SortByColC(wsRecord)

            ResizeRecordColumns(wsRecord)

            For x = 1 To fncGetLastRow(wsRecord, "A")
                wsRecord.Range("B" & CStr(x)).Value = "-->"
            Next
        Catch ex As Exception
            WMsg("RecordNameChanges ERROR " & ex.Message)
        End Try
    End Sub

    Sub ResizeRecordColumns(ByRef wsRecord As Worksheet)
        wsRecord.Range("A1").ColumnWidth = 27.75
        wsRecord.Range("B1").ColumnWidth = 4
        wsRecord.Range("C1").ColumnWidth = 27.75
        wsRecord.Range("D1").ColumnWidth = 14
        wsRecord.Range("E1").ColumnWidth = 12
    End Sub

    Sub DeleteNamesChanged(ByRef wsAdded As Worksheet, ByRef wsRemoved As Worksheet)
        Try
            DeleteRowsWithValues(wsAdded, "G")
            DeleteRowsWithValues(wsRemoved, "G")

            '--Delete column "F"
            wsAdded.Range("F1").EntireColumn.Delete()
            wsRemoved.Range("F1").EntireColumn.Delete()

            SortByColDA(wsAdded)
            SortByColDA(wsRemoved)

        Catch ex As Exception
            WMsg("DeleteNamesChanged ERROR " & ex.Message)
        End Try
    End Sub

    Sub CloseWorkbooks()
        Try
            wbNewPtOld.Close(False)
            wbPanelOld.Close(False)
            wbPanelRecent.Close(False)
            wbNewPtRecent.Close(False)
        Catch ex As Exception
            WMsg("CloseWorkbooks ERROR " & ex.Message)
        End Try
    End Sub

    Sub CloseExcel()
        excelApplication.ScreenUpdating = True
        If Not excelApplication Is Nothing Then
            excelApplication.Quit()
            excelApplication = Nothing
        End If
        GC.Collect()
        GC.WaitForPendingFinalizers()
        GC.Collect()
        GC.WaitForPendingFinalizers()
    End Sub

End Module
