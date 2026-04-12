Attribute VB_Name = "QATMain"
'---------------------------------------------------------------
' SuperQAT - Word QAT Command Dropdown
' Exposes all QAT commands in a single searchable dropdown.
'---------------------------------------------------------------
Option Explicit

Private Const BAR_NAME As String = "SuperQAT"
Private Const CMD_COUNT As Long = 1884

Private cmdIDs() As Long
Private cmdCaptions() As String

Private Sub LoadCommands()
    ReDim cmdIDs(1 To CMD_COUNT)
    ReDim cmdCaptions(1 To CMD_COUNT)
    LoadChunk1 cmdIDs, cmdCaptions
    LoadChunk2 cmdIDs, cmdCaptions
    LoadChunk3 cmdIDs, cmdCaptions
    LoadChunk4 cmdIDs, cmdCaptions
    LoadChunk5 cmdIDs, cmdCaptions
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

    Dim ctrl As Office.CommandBarControl
    Set ctrl = Application.CommandBars.FindControl(ID:=targetID)

    If ctrl Is Nothing Then
        MsgBox "Command " & Chr(39) & cmdCaptions(idx) & Chr(39) & " (ID " & targetID & ") is not available.", vbInformation, "SuperQAT"
        Exit Sub
    End If

    ctrl.Execute
    Exit Sub

ErrHandler:
    MsgBox "Could not run " & Chr(39) & cmdCaptions(idx) & Chr(39) & ": " & Err.Description, vbExclamation, "SuperQAT"
End Sub
