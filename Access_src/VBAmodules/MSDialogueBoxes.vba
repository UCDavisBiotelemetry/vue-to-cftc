Option Compare Database
Option Explicit
'***************** Code Start **************
' This code was originally written by Ken Getz.
' It is not to be altered or distributed, 'except as part of an application.
' You are free to use it in any application,
' provided the copyright notice is left unchanged.
'
' Code originally courtesy of:
' Microsoft Access 95 How-To
' Ken Getz and Paul Litwin
' Waite Group Press, 1996
' Revised to support multiple files:
' 28 December 2007

Type tagOPENFILENAME
    lStructSize As Long
    hwndOwner As Long 'LongPtr
    hInstance As Long 'LongPtr
    strFilter As String
    strCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    strFile As String
    nMaxFile As Long
    strFileTitle As String
    nMaxFileTitle As Long
    strInitialDir As String
    strTitle As String
    Flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    strDefExt As String
    lCustData As Long
    lpfnHook As Long 'LongPtr
    lpTemplateName As String
End Type

Public Type OPENFILENAME
    lStructSize As Long
    hwndOwner As LongPtr
    hInstance As LongPtr
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    Flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As LongPtr
    lpTemplateName As String
End Type

Declare PtrSafe Function aht_apiGetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (OFN As tagOPENFILENAME) As Boolean
Public Declare PtrSafe Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (OFN As OPENFILENAME) As Long
Declare PtrSafe Function aht_apiGetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (OFN As tagOPENFILENAME) As Boolean
Declare PtrSafe Function CommDlgExtendedError Lib "comdlg32.dll" () As Long

Global Const ahtOFN_READONLY = &H1
Global Const ahtOFN_OVERWRITEPROMPT = &H2
Global Const ahtOFN_HIDEREADONLY = &H4
Global Const ahtOFN_NOCHANGEDIR = &H8
Global Const ahtOFN_SHOWHELP = &H10
' You won't use these.
'Global Const ahtOFN_ENABLEHOOK = &H20
'Global Const ahtOFN_ENABLETEMPLATE = &H40
'Global Const ahtOFN_ENABLETEMPLATEHANDLE = &H80
Global Const ahtOFN_NOVALIDATE = &H100
Global Const ahtOFN_ALLOWMULTISELECT = &H200
Global Const ahtOFN_EXTENSIONDIFFERENT = &H400
Global Const ahtOFN_PATHMUSTEXIST = &H800
Global Const ahtOFN_FILEMUSTEXIST = &H1000
Global Const ahtOFN_CREATEPROMPT = &H2000
Global Const ahtOFN_SHAREAWARE = &H4000
Global Const ahtOFN_NOREADONLYRETURN = &H8000
Global Const ahtOFN_NOTESTFILECREATE = &H10000
Global Const ahtOFN_NONETWORKBUTTON = &H20000
Global Const ahtOFN_NOLONGNAMES = &H40000
' New for Windows 95
Global Const ahtOFN_EXPLORER = &H80000
Global Const ahtOFN_NODEREFERENCELINKS = &H100000
Global Const ahtOFN_LONGNAMES = &H200000


Public Function OpenGetFileDialog( _
    Optional ByRef DialogTitle As Variant, _
    Optional ByVal InitialDir As Variant, _
    Optional ByVal Filter As Variant _
) As String
    Dim OpenFile    As OPENFILENAME
    Dim lReturn     As Long

    If IsMissing(DialogTitle) Then DialogTitle = "Default Title"
    If IsMissing(InitialDir) Then InitialDir = "C:\"
    If IsMissing(Filter) Then Filter = ""
    
    OpenFile.lStructSize = LenB(OpenFile)
    OpenFile.hwndOwner = Application.hWndAccessApp
    OpenFile.lpstrFile = String(256, 0)
    OpenFile.nMaxFile = LenB(OpenFile.lpstrFile) - 1
    OpenFile.lpstrFileTitle = OpenFile.lpstrFile
    OpenFile.nMaxFileTitle = OpenFile.nMaxFile
    OpenFile.lpstrInitialDir = InitialDir
    OpenFile.lpstrFilter = Filter
    OpenFile.lpstrTitle = DialogTitle
    OpenFile.Flags = 0
    
    lReturn = GetOpenFileName(OpenFile)
    
    If lReturn = 0 Then
        OpenGetFileDialog = ""
    Else
        OpenGetFileDialog = OpenFile.lpstrFile
    End If
End Function

Function GetOpenFile(Optional varDirectory As Variant, _
    Optional varTitleForDialog As Variant) As Variant

    ' Here's an example that gets an Access database name.
    Dim strFilter As String
    Dim lngFlags As Long
    Dim varFileName As Variant

    ' Specify that the chosen file must already exist,
    ' don't change directories when you're done
    ' Also, don't bother displaying
    ' the read-only box. It'll only confuse people.
    lngFlags = ahtOFN_FILEMUSTEXIST Or _
                ahtOFN_HIDEREADONLY Or ahtOFN_NOCHANGEDIR
    If IsMissing(varDirectory) Then
        varDirectory = ""
    End If
    If IsMissing(varTitleForDialog) Then
        varTitleForDialog = ""
    End If

    ' Define the filter string and allocate space in the "c"
    ' string Duplicate this line with changes as necessary for
    ' more file templates.
    strFilter = ahtAddFilterItem(strFilter, _
                "Access (*.mdb)", "*.MDB;*.MDA")

    ' Now actually call to get the file name.
    varFileName = ahtCommonFileOpenSave( _
                    OpenFile:=True, _
                    InitialDir:=varDirectory, _
                    Filter:=strFilter, _
                    Flags:=lngFlags, _
                    DialogTitle:=varTitleForDialog)
    If Not IsNull(varFileName) Then
        varFileName = TrimNull(varFileName)
    End If
    GetOpenFile = varFileName
End Function

Function ahtCommonFileOpenSave( _
            Optional ByRef Flags As Variant, _
            Optional ByVal InitialDir As Variant, _
            Optional ByVal Filter As Variant, _
            Optional ByVal FilterIndex As Variant, _
            Optional ByVal DefaultExt As Variant, _
            Optional ByVal FileName As Variant, _
            Optional ByVal DialogTitle As Variant, _
            Optional ByVal hwnd As Variant, _
            Optional ByVal OpenFile As Variant) As Variant

    ' This is the entry point you'll use to call the common
    ' file open/save dialog. The parameters are listed
    ' below, and all are optional.
    '
    ' In:
    ' Flags: one or more of the ahtOFN_* constants, OR'd together.
    ' InitialDir: the directory in which to first look
    ' Filter: a set of file filters, set up by calling
    ' AddFilterItem. See examples.
    ' FilterIndex: 1-based integer indicating which filter
    ' set to use, by default (1 if unspecified)
    ' DefaultExt: Extension to use if the user doesn't enter one.
    ' Only useful on file saves.
    ' FileName: Default value for the file name text box.
    ' DialogTitle: Title for the dialog.
    ' hWnd: parent window handle
    ' OpenFile: Boolean(True=Open File/False=Save As)
    ' Out:
    ' Return Value: Either Null or the selected filename
    Dim OFN As tagOPENFILENAME
    Dim strFileName As String
    Dim strFileTitle As String
    Dim fResult As Boolean

    ' Give the dialog a caption title.
    If IsMissing(InitialDir) Then InitialDir = CurDir
    If IsMissing(Filter) Then Filter = ""
    If IsMissing(FilterIndex) Then FilterIndex = 1
    If IsMissing(Flags) Then Flags = 0&
    If IsMissing(DefaultExt) Then DefaultExt = ""
    If IsMissing(FileName) Then FileName = ""
    If IsMissing(DialogTitle) Then DialogTitle = ""
    If IsMissing(hwnd) Then hwnd = Application.hWndAccessApp
    If IsMissing(OpenFile) Then OpenFile = True
    ' Allocate string space for the returned strings.
    strFileName = Left(FileName & String(256, 0), 256)
    strFileTitle = String(256, 0)
    ' Set up the data structure before you call the function
    With OFN
        .lStructSize = Len(OFN)
        .hwndOwner = hwnd
        .strFilter = Filter
        .nFilterIndex = FilterIndex
        .strFile = strFileName
        .nMaxFile = Len(strFileName)
        .strFileTitle = strFileTitle
        .nMaxFileTitle = Len(strFileTitle)
        .strTitle = DialogTitle
        .Flags = Flags
        .strDefExt = DefaultExt
        .strInitialDir = InitialDir
        ' Didn't think most people would want to deal with
        ' these options.
        .hInstance = 0
        '.strCustomFilter = ""
        '.nMaxCustFilter = 0
        .lpfnHook = 0
        'New for NT 4.0
        .strCustomFilter = String(255, 0)
        .nMaxCustFilter = 255
    End With
    ' This will pass the desired data structure to the
    ' Windows API, which will in turn it uses to display
    ' the Open/Save As Dialog.
    If OpenFile Then
        fResult = aht_apiGetOpenFileName(OFN)
    Else
        fResult = aht_apiGetSaveFileName(OFN)
    End If

    ' The function call filled in the strFileTitle member
    ' of the structure. You'll have to write special code
    ' to retrieve that if you're interested.
    If fResult Then
        ' You might care to check the Flags member of the
        ' structure to get information about the chosen file.
        ' In this example, if you bothered to pass in a
        ' value for Flags, we'll fill it in with the outgoing
        ' Flags value.
        If Not IsMissing(Flags) Then Flags = OFN.Flags
        If Flags And ahtOFN_ALLOWMULTISELECT Then
            ' Return the full array.
            Dim items As Variant
            Dim value As String
            value = OFN.strFile
            ' Get rid of empty items:
            Dim i As Integer
            For i = Len(value) To 1 Step -1
              If Mid$(value, i, 1) <> Chr$(0) Then
                Exit For
              End If
            Next i
            value = Mid(value, 1, i)

            ' Break the list up at null characters:
            items = Split(value, Chr(0))

            ' Loop through the items in the "array",
            ' and build full file names:
            Dim numItems As Integer
            Dim result() As String

            numItems = UBound(items) + 1
            If numItems > 1 Then
                ReDim result(0 To numItems - 2)
                For i = 1 To numItems - 1
                    result(i - 1) = FixPath(items(0)) & items(i)
                Next i
                ahtCommonFileOpenSave = result
            Else
                ' If you only select a single item,
                ' Windows just places it in item 0.
                ahtCommonFileOpenSave = items(0)
            End If
        Else
            ahtCommonFileOpenSave = TrimNull(OFN.strFile)
        End If
    Else
        ahtCommonFileOpenSave = vbNullString
    End If
End Function

Function ahtAddFilterItem(strFilter As String, _
    strDescription As String, Optional varItem As Variant) As String

    ' Tack a new chunk onto the file filter.
    ' That is, take the old value, stick onto it the description,
    ' (like "Databases"), a null character, the skeleton
    ' (like "*.mdb;*.mda") and a final null character.

    If IsMissing(varItem) Then varItem = "*.*"
    ahtAddFilterItem = strFilter & _
                strDescription & vbNullChar & _
                varItem & vbNullChar
End Function

Private Function TrimNull(ByVal strItem As String) As String
    Dim intPos As Integer

    intPos = InStr(strItem, vbNullChar)
    If intPos > 0 Then
        TrimNull = Left(strItem, intPos - 1)
    Else
        TrimNull = strItem
    End If
End Function

Private Function FixPath(ByVal path As String) As String
    If Right$(path, 1) <> "\" Then
        FixPath = path & "\"
    Else
        FixPath = path
    End If
End Function

'************** Code End *****************


'Import from CSV
Function AcDialogSelectFile() As String
 'http://stackoverflow.com/questions/1091484/how-to-show-open-file-dialog-in-access-2007-vba
 Dim f    As Object
 Set f = Application.FileDialog(3)
 f.AllowMultiSelect = True
 f.Show
End Function

Function CbrShortCutMenus(Optional SearchFor As String)

    Dim cbr As Object
    Dim intCount As Integer
    Dim outStr As String
    outStr = "list is"
    For Each cbr In Application.CommandBars
        If cbr.Type = 2 Then
            If SearchFor = "" Then SearchFor = " "
            If InStr(cbr.Name & " ", SearchFor) > 0 Then
                Debug.Print cbr.Name
                outStr = outStr & vbCrLf & cbr.Name
            End If
        End If
    Next
    MsgBox (outStr)
    CommandBars("Database Table/Query").ShowPopup
    CbrShortCutMenus = outStr
End Function

Function BuildCmdBar()
    Dim obj As AccessObject
    Dim dbs As Object
    Dim ctrlListBox As ListBox
    Set dbs = Application.CurrentData
    Dim TBar As CommandBar
    Dim NewDD As CommandBarControl
    Dim ActStr As String
    CommandBars("TableList").Delete
    Set TBar = CommandBars.Add("TableList", msoBarPopup, False, False)
    Set NewDD = TBar.Controls.Add(Type:=msoControlButton)
    NewDD.Caption = "Choose Table From List"
    NewDD.Tag = "NameButtonInContextMenu"
    NewDD.BeginGroup = True
    NewDD.Style = msoButtonAutomatic
    NewDD.FaceId = 487
    For Each obj In dbs.AllTables
        If (InStr(1, obj.Name, "MSys") = False) Then
            Set NewDD = CommandBars("TableList").Controls.Add(Type:=msoControlButton)
            With NewDD
                If .Index = 2 Then .BeginGroup = True
                .Style = msoButtonCaption
                .Caption = obj.Name
                .Tag = obj.Name
                .Parameter = obj.Name
   '             ActStr = "grabSelection(" & NewDD.Index & ")"
                ActStr = "=grabSelection(" & NewDD.Index & ")"
                .OnAction = ActStr
            End With
        End If
    Next obj
End Function
    
Function grabSelection(Optional ByVal i As Integer = 1) As String
'    MsgBox (CommandBars("TableList").Controls(i).Parameter & " Selected")
    retval = (CommandBars("TableList").Controls(i).Parameter)
    grabSelection = (CommandBars("TableList").Controls(i).Parameter)
 '   CommandBars("TableList").Controls(i).Caption
End Function
Function Test(Arg1 As String)
    MsgBox (Arg1)
End Function
