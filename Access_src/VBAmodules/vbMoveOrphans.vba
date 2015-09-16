Option Compare Database
Option Explicit
Global retval As String

Function MoveOrphans(Optional ByVal tabl As String = "")
    Dim stmt As String
    Dim stmt2 As String
    Dim boxval
    Dim db As DAO.Database
    If tabl = "" Or tabl = Null Or tabl = " " Then
        MsgBox ("No Table Selected")
        Exit Function
    End If
    Set db = CurrentDb
    stmt = "INSERT INTO DetectOrphans SELECT * FROM " & tabl & " AS Orphans WHERE NOT EXISTS (SELECT 1 FROM Import_Deployments AS Dep WHERE Orphans.VR2SN = Dep.VR2SN AND Orphans.DetectDate BETWEEN Dep.[Start] AND Dep.[Stop]);"
    If MsgBox("About to execute " & vbCrLf & stmt, vbOKCancel) = vbCancel Then Exit Function
    On Error GoTo InsertError
        With db
            .Execute (stmt) 'Try to append to existing table
            MsgBox ("Number of Records Affected: [" & .RecordsAffected & "]")
        End With
    'on error resume next 'for the first statement would result in the deletion occuring even if the insert failed.
    On Error Resume Next
        stmt = "DELETE FROM " & tabl & " AS Orphans WHERE NOT EXISTS (SELECT 1 FROM Import_Deployments AS Dep WHERE Orphans.VR2SN = Dep.VR2SN AND Orphans.DetectDate BETWEEN Dep.[Start] AND Dep.[Stop]);"
        boxval = MsgBox("About to delete records from source table via " & vbCrLf & stmt, vbOKCancel)
        If boxval = vbCancel Then Exit Function
    On Error GoTo DeleteError
        With db
            .Execute (stmt) 'remove from old table
            MsgBox ("Number of Records Affected: [" & .RecordsAffected & "]")
        End With
Exit Function
InsertError:
'    stmt = "SELECT * INTO DetectOrphans FROM " & tabl & " AS Orphans WHERE NOT EXISTS (SELECT 1 FROM Deployments AS Dep WHERE Orphans.VR2SN = Dep.VR2SN AND Orphans.DetectDate BETWEEN Dep.[Start] AND Dep.[Stop]);"
    stmt2 = "SELECT * INTO DetectOrphans FROM " & tabl & " WHERE 1=0;"
'orig'    stmt = "INSERT INTO DetectOrphans SELECT * FROM " & tabl & " AS Orphans WHERE NOT EXISTS (SELECT 1 FROM Deployments AS Dep WHERE Orphans.VR2SN = Dep.VR2SN AND Orphans.DetectDate BETWEEN Dep.[Start] AND Dep.[Stop]);"
    boxval = MsgBox("Non-Critical Error was " & Err.Number & " " & Err.Description & vbCrLf & "Keep in mind that this is likely because the DetectOrphans table may not yet exist" & vbCrLf & "The next step will be to attempt to create the table via" & vbCrLf & stmt2 & vbCrLf & "followed by the original attempted SQL statement", vbOKCancel)
    If boxval = vbCancel Then Exit Function
    Err.Clear
    On Error GoTo Getmeoutofhere
        With db
            .Execute (stmt2) 'Create New Table
            .Execute (stmt) 'Fill
' /*            MsgBox ("Number of Records Affected: [" & .RecordsAffected & "]")*/
        End With
    Resume Next
Getmeoutofhere:
    MsgBox ("Final error " & Err.Number & " " & Err.Description & "...exiting")
    Exit Function
DeleteError:
    MsgBox ("Error deleting " & Err.Number & " " & Err.Description)
    Exit Function
End Function

'Sub AssignIt()
 '   With Application.CommandBars("Cust1").Controls(1)
  '      .OnAction = "Test(" & Chr(34) & "First line" & Chr(34) & "," & Chr(34) & "Second Line" & Chr(34) & ")"
   ' End With
'End Sub
Function UseCommandBar(Optional cappy As String = "Select Detections Table From List") As String
    Dim ActStr As String
    Dim b As CommandBarControl
    Dim i As Integer
    retval = ""
    CommandBars("TableList").Controls(1).Caption = cappy
    CommandBars("TableList").ShowPopup
'with c("TL")
'        .Enabled = True
'        .Visible = True
'        For Each b In CommandBars("TableList").Controls
'            ActStr = "=grabSelection(" & b.Index & ")"
'             b.OnAction = ActStr
'        Next b
'    End With
    UseCommandBar = retval
End Function


