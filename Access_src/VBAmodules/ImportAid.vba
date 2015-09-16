Option Compare Database
Option Explicit
Function imptDetections()
    On Error Resume Next
    Dim txtConnection, txtRecordset ' As ADODB.Connection As ADODB.Recordset
    Dim inFileName As String, outTableName As String, stmt As String
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set txtConnection = CreateObject("ADODB.Connection")
    Set txtRecordset = CreateObject("ADODB.Recordset")
    Dim outR, tR, oCmd, tConnect
    Set outR = CreateObject("ADODB.Recordset")
    Set tR = CreateObject("ADODB.Recordset")
    Set tConnect = CurrentProject.Connection
    Dim sFilt As String
    Dim ddFilt As String
    Dim bNoMatch As Boolean
    Dim rRow As Variant
    Dim i As Long, t As Integer
    Dim sQry As String
    Dim pth As String, fn As String
    Const adOpenStatic = 3
    Const adLockOptimistic = 3
    Const adCmdText = &H1
    Dim DD As Date, BDD As Date
    DoCmd.Hourglass True
    inFileName = MsDialogSelectFile("Detections")
    outTableName = UseCommandBar("Table to Import Detections Into")
    If ((inFileName = "" Or inFileName = Null Or inFileName = " ") Or (outTableName = "" Or outTableName = Null Or outTableName = " ")) Then
        Set txtConnection = Nothing
        Set txtRecordset = Nothing
        Set outR = Nothing
        Set fso = Nothing
        MsgBox ("Required Parameter Not Selected")
        Exit Function
    End If
        
    pth = fso.GetParentFolderName(inFileName)
    fn = fso.GetFileName(inFileName)
    
    Call CreateSchemaFile(False, pth & "\", fn, "Import_Detections")
'    txtConnection.Open "Provider=Microsoft.ACE.OLEDB.12.0;" & _
'        "Data Source='" & pth & "';" & _
'        "Extended Properties=""text;HDR=NO;FMT=Delimited"""

    If MsgBox("About to import into [" & outTableName & "] from file """ & inFileName & """", vbOKCancel) = vbCancel Then Exit Function
        
'rather than currentproject.connection.execute, could use DAO which should be faster than ADODB. However, when dealing with the text file, we "need" ADODB
'On Error Resume Next
On Error GoTo err_temp_table_create
    sQry = "CREATE TABLE [tmp_" & outTableName & "] (" & _
        "[TagID] int NOT NULL, [Codespace] varchar(25) NOT NULL, [DetectDate] datetime NOT NULL, " & _
        "[VR2SN] int NOT NULL, [Data] float NULL, [Units] varchar(50) NULL, [Data2] float NULL, [Units2] varchar(50) NULL, [BasicDD] datetime NOT NULL, [DifDD] float NULL, " & _
        "CONSTRAINT [PK_DT] PRIMARY KEY ([TagID], [Codespace], [DetectDate], [VR2SN]))"
    CurrentDb.Execute sQry
    sQry = "CREATE INDEX iBasicDD ON [tmp_" & outTableName & "] ([TagID], [Codespace], [BasicDD], [VR2SN]) WITH DISALLOW NULL"
    CurrentDb.Execute sQry
'On Error GoTo 0
'    CurrentProject.Connection.Execute "INSERT INTO [tmp_" & outTableName & "] (SELECT *, CDate(CStr(DetectDate)) AS BasicDD FROM [" & outTableName & "]"
On Error GoTo err_clone_dets_to_temp
    sQry = "INSERT INTO [tmp_" & outTableName & "] ([TagID], [Codespace], [DetectDate], [VR2SN], [Data], [Units], [Data2], [Units2], [BasicDD], [DifDD]) " & _
        "SELECT TagID, Codespace, DetectDate, VR2SN, Data, Units, Data2, Units2, (CDate(CStr(DetectDate))) as BasicDD, (CDbl(DetectDate)-CDbl(CDate(CStr(DetectDate)))) as DifDD FROM [" & outTableName & "]"
    Debug.Print "Executing " & sQry
    CurrentDb.Execute sQry
'    CurrentProject.Connection.CursorLocation = adUseClient
'        sQry = "SELECT * FROM [tmp_" & outTableName & "] WHERE 1=0"
'        tR.Open sQry, CurrentProject.Connection, adOpenForwardOnly, adLockOptimistic, adCmdText
    tConnect.BeginTrans
On Error GoTo err_tempfill_from_file
        '"Provider=Microsoft.ACE.OLEDB.12.0;"
        txtConnection.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
            "Data Source=" & pth & "\;" & _
            "Extended Properties=""text;HDR=NO;FMT=Delimited"""
        txtConnection.Open
        txtRecordset.Open "SELECT * FROM [" & fn & "]", txtConnection, adOpenForwardOnly, adLockReadOnly, adCmdText 'txtopen
        sQry = "SELECT * FROM [tmp_" & outTableName & "] WHERE 1=0"
        Debug.Print "Opened temp; Executing " & sQry & " and opening ADO recordset"
        tR.Open sQry, tConnect, adOpenForwardOnly, adLockOptimistic, adCmdText
        Debug.Print "file opened if no error printed immediately before"
        i = 0
        txtRecordset.MoveFirst
        Do Until txtRecordset.EOF
          With txtRecordset.Fields
            If (Not (IsNull(.Item("TagID")))) Then
                DD = CDateMsec(.Item("DetectDate"))
                BDD = CDate(CStr(DD))
                i = i + 1
                tR.AddNew
                tR!TagID = .Item("TagID")
                tR!Codespace = .Item("Codespace")
                tR!DetectDate = DD
                tR!VR2SN = .Item("VR2SN")
                tR!Data = .Item("Data")
                tR!Units = .Item("Units")
                tR!Data2 = .Item("Data2")
                tR!Units2 = .Item("Units2")
                tR!BasicDD = BDD
                tR!DifDD = (CDbl(DD) - CDbl(BDD))
                tR.Update
            End If
          End With
          txtRecordset.MoveNext
        Loop
        If Not (txtRecordset Is Nothing) Then
            If (txtRecordset.State And adStateOpen) = adStateOpen Then txtRecordset.Close
            Set txtRecordset = Nothing
        End If
        If Not (txtConnection Is Nothing) Then
            If (txtConnection.State And adStateOpen) = adStateOpen Then txtConnection.Close
            Set txtConnection = Nothing
        End If
        If Not (tR Is Nothing) Then
            If (tR.State And adStateOpen) = adStateOpen Then tR.Close
            Set tR = Nothing
        End If
    tConnect.CommitTrans
    DBEngine.BeginTrans
On Error GoTo rb
        sQry = "DELETE FROM [" & outTableName & "]"
        CurrentDb.Execute sQry
        DBEngine.Idle dbRefreshCache
        sQry = "INSERT INTO [" & outTableName & "] SELECT T.[TagID], T.[Codespace], T.[DetectDate], T.[VR2SN], T.[Data], T.[Units], T.[Data2], T.[Units2] " & _
            "FROM [tmp_" & outTableName & "] AS T WHERE ABS(T.[DifDD]) IN (SELECT MAX(ABS(T2.[DifDD])) FROM [tmp_" & outTableName & "] AS T2 GROUP BY T2.[TagID], T2.[Codespace], T2.[BasicDD], T2.[VR2SN])"
        Debug.Print sQry
        CurrentDb.Execute sQry ', CurrentProject.Connection, adOpenStatic, adLockBatchOptimistic, adCmdText
        Debug.Print CurrentDb.RecordsAffected & " records affected"
    DBEngine.CommitTrans
    On Error Resume Next
    DBEngine.Idle dbRefreshCache
    CurrentDb.Execute "DROP TABLE [tmp_" & outTableName & "]"
ext:
    On Error GoTo 0
'    txtRecordset.Close
    If Not (txtRecordset Is Nothing) Then
        If (txtRecordset.State And adStateOpen) = adStateOpen Then txtRecordset.Close
        Set txtRecordset = Nothing
    End If
    If Not (txtConnection Is Nothing) Then
        If (txtConnection.State And adStateOpen) = adStateOpen Then txtConnection.Close
        Set txtConnection = Nothing
    End If
    If Not (outR Is Nothing) Then
        If (outR.State And adStateOpen) = adStateOpen Then outR.Close
        Set outR = Nothing
    End If
    If Not (tR Is Nothing) Then
        If (tR.State And adStateOpen) = adStateOpen Then tR.Close
        Set tR = Nothing
    End If
    Set fso = Nothing
    DoCmd.Hourglass False
    Exit Function
rb:
    Debug.Print Err.Description & CStr(Err.Number)
    DBEngine.Rollback
'    Err.Raise
    GoTo ext
err_temp_table_create:
    Debug.Print "error with temp table creation: " & Err.Description & CStr(Err.Number)
    Resume Next
err_clone_dets_to_temp:
    Debug.Print "error cloning detections to temp table: " & Err.Description & CStr(Err.Number)
    Resume Next
err_tempfill_from_file:
    Debug.Print "error with filling temp table from file: " & Err.Description & CStr(Err.Number)
    Debug.Print "dir: " & pth & "   file: " & fn & "   table:" & outTableName & "   iter: " & i
    Err.Clear
    On Error GoTo err_temp_rb
    Resume Next
err_temp_rb:
    Debug.Print Err.Description & CStr(Err.Number)
    tConnect.Rollback
    GoTo ext
End Function


Function MsDialogSelectFile(Optional tabType As String = "Detections") As String 'Variant
    Dim strFilter As String
    Dim lngFlags As Long
    strFilter = ahtAddFilterItem(strFilter, "CSV Text Files (*.csv)", "*.CSV")
    strFilter = ahtAddFilterItem(strFilter, "Other Text Files (*.txt)", "*.TXT;*.TSV")
'    strFilter = ahtAddFilterItem(strFilter, "Excel Files (*.xls(x))", "*.XLS*")
    strFilter = ahtAddFilterItem(strFilter, "All Files (*.*)", "*.*")

    ' Uncomment this line to try the example
    ' allowing multiple file names:
    ' lngFlags = ahtOFN_ALLOWMULTISELECT Or ahtOFN_EXPLORER

    Dim result As Variant

    result = ahtCommonFileOpenSave(Filter:=strFilter, FilterIndex:=1, Flags:=lngFlags, DialogTitle:="Please Select Your " & tabType & " File")
    MsDialogSelectFile = result
End Function


Public Function CreateSchemaFile(bIncFldNames As Boolean, _
                                       sPath As String, _
                                       sSectionName As String, _
                                       sTblQryName As String) As Boolean
' http://support.microsoft.com/kb/155512
'     Parameter value
'   ------------------------------------------------------------------------
'   bIncFldNames     True/False, stating if the first row of the text file
'                    has column names
'   sPath            Full path to the folder where Schema.ini will reside
'   sSectionName     Schema.ini section name; must be the same as the name
'                    of the text file it describes
'   sTblQryName      Name of the table or query for which you want to
'                    create a Schema.ini file    Dim strFilter As String
         
         Dim Msg As String ' For error handling.
         On Local Error GoTo CreateSchemaFile_Err
         Dim ws As Workspace, db As Database
         Dim tblDef As DAO.TableDef, fldDef As DAO.Field
         Dim i As Integer, Handle As Integer
         Dim fldName As String, fldDataInfo As String
         ' -----------------------------------------------
         ' Set DAO objects.
         ' -----------------------------------------------
         Set db = CurrentDb()
         ' -----------------------------------------------
         ' Open schema file for append.
         ' -----------------------------------------------
         Handle = FreeFile
         Open sPath & "schema.ini" For Output Access Write As #Handle
         ' -----------------------------------------------
         ' Write schema header.
         ' -----------------------------------------------
         Print #Handle, "[" & sSectionName & "]"
         Print #Handle, "ColNameHeader = " & _
                         IIf(bIncFldNames, "True", "False")
         Print #Handle, "CharacterSet = ANSI"
         Print #Handle, "Format = CSVDelimited"
         ' -----------------------------------------------
         ' Get data concerning schema file.
         ' -----------------------------------------------
         Set tblDef = db.TableDefs(sTblQryName)
         With tblDef
            For i = 0 To .Fields.Count - 1
               Set fldDef = tblDef.Fields(i)
               With fldDef
                  fldName = .Name
                  Select Case .Type
                     Case dbBoolean
                        fldDataInfo = "Bit"
                     Case dbByte
                        fldDataInfo = "Byte"
                     Case dbInteger
                        fldDataInfo = "Short"
                     Case dbLong
                        fldDataInfo = "Integer"
                     Case dbCurrency
                        fldDataInfo = "Currency"
                     Case dbSingle
                        fldDataInfo = "Single"
                     Case dbDouble
                        fldDataInfo = "Double"
                     Case dbDate
                        fldDataInfo = "Char Width 30" '"Date" was used, but does not allow for miliseconds
                     Case dbText
                        fldDataInfo = "Char Width " & .Size
                     Case dbLongBinary
                        fldDataInfo = "OLE"
                     Case dbMemo
                        fldDataInfo = "LongChar"
                     Case dbGUID
                        fldDataInfo = "Char Width 16"
                  End Select
                  Print #Handle, "Col" & Format$(i + 1) _
                                  & "=" & fldName & Space$(1) _
                                  & fldDataInfo
               End With
            Next i
         End With
'         MsgBox sPath & "SCHEMA.INI has been created."
         CreateSchemaFile = True
CreateSchemaFile_End:
         Close Handle
         Exit Function
CreateSchemaFile_Err:
         Msg = "Error #: " & Format$(Err.Number) & vbCrLf
         Msg = Msg & Err.Description
         MsgBox Msg
         Resume CreateSchemaFile_End
End Function

