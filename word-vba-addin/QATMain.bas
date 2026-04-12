Attribute VB_Name = "QATMain"
'---------------------------------------------------------------
' SuperQAT - Word QAT Command Dropdown
' Exposes all QAT commands in a single searchable dropdown.
'---------------------------------------------------------------
Option Explicit

Private Const BAR_NAME As String = "SuperQAT"
Private Const CMD_COUNT As Long = 1067

Private cmdIDs() As Long
Private cmdCaptions() As String

Private Sub LoadCommands()
    ReDim cmdIDs(1 To CMD_COUNT)
    ReDim cmdCaptions(1 To CMD_COUNT)
    LoadChunk1 cmdIDs, cmdCaptions
    LoadChunk2 cmdIDs, cmdCaptions
    LoadChunk3 cmdIDs, cmdCaptions
End Sub

Public Sub AutoExec()
    On Error Resume Next
    Application.CommandBars(BAR_NAME).Delete
    On Error GoTo 0

    LoadCommands

    Dim bar As Office.CommandBar
    Set bar = Application.CommandBars.Add(Name:=BAR_NAME, Position:=msoBarTop, Temporary:=True)
    bar.Visible = True

    Dim cbo As Office.CommandBarComboBox
    Set cbo = bar.Controls.Add(Type:=msoControlComboBox)
    With cbo
        .Caption = "QAT Command"
        .Tag = "SuperQAT_Combo"
        .Style = msoComboLabel
        .Width = 300
        .OnAction = "RunSelectedCommand"
        Dim i As Long
        For i = 1 To CMD_COUNT
            .AddItem cmdCaptions(i)
        Next i
    End With
End Sub

Public Sub AutoExit()
    On Error Resume Next
    Application.CommandBars(BAR_NAME).Delete
    On Error GoTo 0
End Sub

Public Sub RunSelectedCommand()
    On Error GoTo ErrHandler

    Dim idx As Long
    Dim cb As Office.CommandBarComboBox
    Set cb = Application.CommandBars(BAR_NAME).FindControl(Tag:="SuperQAT_Combo")
    If cb Is Nothing Then Exit Sub

    idx = cb.ListIndex
    If idx < 1 Or idx > CMD_COUNT Then Exit Sub

    If (Not Not cmdIDs) = 0 Then LoadCommands

    Dim targetID As Long
    targetID = cmdIDs(idx)

    ' Strategy 1: Find an existing CommandBar control and execute it
    Dim ctrl As Office.CommandBarControl
    Set ctrl = Application.CommandBars.FindControl(ID:=targetID)

    If Not ctrl Is Nothing Then
        ctrl.Execute
        Exit Sub
    End If

    ' Strategy 2: Create a temporary control by ID, execute, then delete.
    ' Many commands exist as valid IDs but have no loaded CommandBar control
    ' in modern Word. Adding one by ID instantiates it so we can run it.
    Dim tmpBar As Office.CommandBar
    Dim tmpCtrl As Office.CommandBarControl

    On Error Resume Next
    Set tmpBar = Application.CommandBars.Add("SuperQAT_Tmp", msoBarPopup, , True)
    Set tmpCtrl = tmpBar.Controls.Add(ID:=targetID)
    On Error GoTo ErrHandler

    If tmpCtrl Is Nothing Then
        On Error Resume Next
        tmpBar.Delete
        On Error GoTo 0
        MsgBox "Command " & Chr(39) & cmdCaptions(idx) & Chr(39) & " (ID " & targetID & ") is not available in this version of Word.", vbInformation, "SuperQAT"
        Exit Sub
    End If

    tmpCtrl.Execute

    On Error Resume Next
    tmpBar.Delete
    On Error GoTo 0
    Exit Sub

ErrHandler:
    On Error Resume Next
    If Not tmpBar Is Nothing Then tmpBar.Delete
    On Error GoTo 0
    MsgBox "Could not run " & Chr(39) & cmdCaptions(idx) & Chr(39) & ": " & Err.Description, vbExclamation, "SuperQAT"
End Sub
