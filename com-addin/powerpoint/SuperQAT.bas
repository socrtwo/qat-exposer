Attribute VB_Name = "SuperQAT"
'==============================================================================
' SuperQAT v3.0.0 - PowerPoint COM Add-in
' Executes ALL 769 PowerPoint ribbon commands via CommandBars.ExecuteMso
'
' INSTALL:
'   1. Open PowerPoint, press Alt+F11 to open VBA Editor
'   2. File > Import File > select this .bas file
'   3. File > Import File > select SuperQATData.bas
'   4. File > Import File > select SuperQATForm.frm
'   5. Close VBA Editor
'   6. File > Save As > PowerPoint Add-In (.ppam)
'      Save to: %appdata%\Microsoft\AddIns\SuperQAT.ppam
'   7. File > Options > Add-Ins > Manage: PowerPoint Add-ins > Go
'      Check "SuperQAT" and click OK.
'==============================================================================
Option Explicit

Public Const CMD_COUNT As Long = 769

Private mCommands() As String

Public Sub InitCommands()
    If Not Not mCommands Then Exit Sub
    ReDim mCommands(0 To CMD_COUNT - 1, 0 To 1)
    Dim i As Long
    Dim raw As Variant
    raw = GetAllCommands()
    For i = 0 To UBound(raw)
        mCommands(i, 0) = raw(i)(0)
        mCommands(i, 1) = raw(i)(1)
    Next i
End Sub

Public Sub ExecuteCommand(ByVal idMso As String)
    On Error GoTo ErrHandler
    Application.CommandBars.ExecuteMso idMso
    Exit Sub
ErrHandler:
    MsgBox "Could not execute '" & idMso & "':" & vbCrLf & Err.Description, _
           vbExclamation, "SuperQAT"
End Sub

Public Sub ShowSuperQAT()
    InitCommands
    Dim frm As New SuperQATForm
    frm.LoadCommands mCommands, CMD_COUNT
    frm.Show vbModeless
End Sub

Public Sub Auto_Open()
    On Error Resume Next
    Dim ctl As CommandBarControl
    For Each ctl In Application.CommandBars("Menu Bar").Controls
        If ctl.Tag = "SuperQAT" Then ctl.Delete
    Next ctl

    Dim btn As CommandBarButton
    Set btn = Application.CommandBars("Menu Bar").Controls.Add( _
        Type:=msoControlButton, Temporary:=True)
    btn.Caption = "SuperQAT"
    btn.Tag = "SuperQAT"
    btn.Style = msoButtonCaption
    btn.OnAction = "ShowSuperQAT"
    On Error GoTo 0
End Sub

Public Function GetCommandId(idx As Long) As String
    GetCommandId = mCommands(idx, 0)
End Function

Public Function GetCommandLabel(idx As Long) As String
    GetCommandLabel = mCommands(idx, 1)
End Function
