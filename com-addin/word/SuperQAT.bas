Attribute VB_Name = "SuperQAT"
'==============================================================================
' SuperQAT v3.0.0 - Word COM Add-in
' Executes ALL 1,334 Word ribbon commands via CommandBars.ExecuteMso
'
' INSTALL:
'   1. Open Word, press Alt+F11 to open VBA Editor
'   2. File > Import File > select this .bas file
'   3. File > Import File > select SuperQATForm.frm
'   4. Close VBA Editor
'   5. File > Save As > Word Macro-Enabled Template (.dotm)
'      Save to: %appdata%\Microsoft\Word\STARTUP\SuperQAT.dotm
'   6. Restart Word. SuperQAT appears on the Add-ins tab.
'==============================================================================
Option Explicit

' Number of commands available
Public Const CMD_COUNT As Long = 1334

' Store commands as a module-level array
Private mCommands() As String  ' (0 To CMD_COUNT-1, 0 To 1) = idMso, Label

'------------------------------------------------------------------------------
' Initialize the command list
'------------------------------------------------------------------------------
Public Sub InitCommands()
    If Not Not mCommands Then Exit Sub  ' already initialized
    ReDim mCommands(0 To CMD_COUNT - 1, 0 To 1)
    Dim i As Long
    Dim raw As Variant
    raw = GetAllCommands()
    For i = 0 To UBound(raw)
        mCommands(i, 0) = raw(i)(0)  ' idMso
        mCommands(i, 1) = raw(i)(1)  ' label
    Next i
End Sub

'------------------------------------------------------------------------------
' Execute a command by its idMso name
'------------------------------------------------------------------------------
Public Sub ExecuteCommand(ByVal idMso As String)
    On Error GoTo ErrHandler
    Application.CommandBars.ExecuteMso idMso
    Exit Sub
ErrHandler:
    MsgBox "Could not execute '" & idMso & "':" & vbCrLf & Err.Description, _
           vbExclamation, "SuperQAT"
End Sub

'------------------------------------------------------------------------------
' Show the SuperQAT command picker
'------------------------------------------------------------------------------
Public Sub ShowSuperQAT()
    InitCommands
    Dim frm As New SuperQATForm
    frm.LoadCommands mCommands, CMD_COUNT
    frm.Show vbModeless
End Sub

'------------------------------------------------------------------------------
' Auto-open: add a toolbar button when the template loads
'------------------------------------------------------------------------------
Public Sub AutoExec()
    ' Add menu item to Add-ins tab
    On Error Resume Next
    Dim cb As CommandBar
    Set cb = Application.CommandBars("Menu Bar")

    ' Remove old one if exists
    Dim ctl As CommandBarControl
    For Each ctl In cb.Controls
        If ctl.Tag = "SuperQAT" Then ctl.Delete
    Next ctl

    ' Add new button
    Dim btn As CommandBarButton
    Set btn = cb.Controls.Add(Type:=msoControlButton, Temporary:=True)
    btn.Caption = "SuperQAT"
    btn.Tag = "SuperQAT"
    btn.Style = msoButtonCaption
    btn.OnAction = "ShowSuperQAT"
    On Error GoTo 0
End Sub

'------------------------------------------------------------------------------
' Get command count and command at index
'------------------------------------------------------------------------------
Public Function GetCommandId(idx As Long) As String
    GetCommandId = mCommands(idx, 0)
End Function

Public Function GetCommandLabel(idx As Long) As String
    GetCommandLabel = mCommands(idx, 1)
End Function
