VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SuperQATForm
   Caption         =   "SuperQAT - Word Commands"
   ClientHeight    =   8400
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4800
   OleObjectBlob   =   "SuperQATForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SuperQATForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==============================================================================
' SuperQAT UserForm - Command Picker with Search
'==============================================================================
Option Explicit

Private mCommands() As String
Private mCount As Long
Private mFiltered() As Long  ' indices into mCommands
Private mFilteredCount As Long

'------------------------------------------------------------------------------
' Initialize the form controls programmatically
'------------------------------------------------------------------------------
Private Sub UserForm_Initialize()
    Me.Width = 360
    Me.Height = 520

    ' Search box
    Dim lblSearch As MSForms.Label
    Set lblSearch = Me.Controls.Add("Forms.Label.1", "lblSearch")
    lblSearch.Caption = "Search:"
    lblSearch.Left = 6: lblSearch.Top = 6: lblSearch.Width = 42: lblSearch.Height = 16

    Dim txtSearch As MSForms.TextBox
    Set txtSearch = Me.Controls.Add("Forms.TextBox.1", "txtSearch")
    txtSearch.Left = 48: txtSearch.Top = 4: txtSearch.Width = 290: txtSearch.Height = 20

    ' Command list
    Dim lstCommands As MSForms.ListBox
    Set lstCommands = Me.Controls.Add("Forms.ListBox.1", "lstCommands")
    lstCommands.Left = 6: lstCommands.Top = 28
    lstCommands.Width = 332: lstCommands.Height = 420

    ' Run button
    Dim btnRun As MSForms.CommandButton
    Set btnRun = Me.Controls.Add("Forms.CommandButton.1", "btnRun")
    btnRun.Caption = "Run Command"
    btnRun.Left = 6: btnRun.Top = 454
    btnRun.Width = 160: btnRun.Height = 28

    ' Count label
    Dim lblCount As MSForms.Label
    Set lblCount = Me.Controls.Add("Forms.Label.1", "lblCount")
    lblCount.Caption = ""
    lblCount.Left = 172: lblCount.Top = 460: lblCount.Width = 166: lblCount.Height = 16
    lblCount.TextAlign = fmTextAlignRight
End Sub

'------------------------------------------------------------------------------
' Load commands from the module
'------------------------------------------------------------------------------
Public Sub LoadCommands(cmds() As String, count As Long)
    mCount = count
    mCommands = cmds
    ReDim mFiltered(0 To mCount - 1)
    FilterCommands ""
End Sub

'------------------------------------------------------------------------------
' Filter and display commands
'------------------------------------------------------------------------------
Private Sub FilterCommands(query As String)
    Dim lst As MSForms.ListBox
    Set lst = Me.Controls("lstCommands")
    lst.Clear

    mFilteredCount = 0
    Dim i As Long
    Dim q As String
    q = LCase(query)

    For i = 0 To mCount - 1
        If Len(q) = 0 Or InStr(1, LCase(mCommands(i, 1)), q) > 0 Or _
           InStr(1, LCase(mCommands(i, 0)), q) > 0 Then
            mFiltered(mFilteredCount) = i
            mFilteredCount = mFilteredCount + 1
            lst.AddItem mCommands(i, 1)
        End If
    Next i

    Me.Controls("lblCount").Caption = mFilteredCount & " of " & mCount & " commands"
End Sub

'------------------------------------------------------------------------------
' Search box change handler
'------------------------------------------------------------------------------
Private Sub txtSearch_Change()
    FilterCommands Me.Controls("txtSearch").Text
End Sub

'------------------------------------------------------------------------------
' Run button click
'------------------------------------------------------------------------------
Private Sub btnRun_Click()
    RunSelected
End Sub

'------------------------------------------------------------------------------
' Double-click to run
'------------------------------------------------------------------------------
Private Sub lstCommands_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    RunSelected
End Sub

'------------------------------------------------------------------------------
' Execute the selected command
'------------------------------------------------------------------------------
Private Sub RunSelected()
    Dim lst As MSForms.ListBox
    Set lst = Me.Controls("lstCommands")

    If lst.ListIndex < 0 Then
        MsgBox "Pick a command first.", vbInformation, "SuperQAT"
        Exit Sub
    End If

    Dim cmdIdx As Long
    cmdIdx = mFiltered(lst.ListIndex)

    Dim idMso As String
    idMso = mCommands(cmdIdx, 0)

    On Error GoTo ErrHandler
    Application.CommandBars.ExecuteMso idMso
    Exit Sub

ErrHandler:
    MsgBox "Could not execute '" & idMso & "':" & vbCrLf & _
           Err.Description, vbExclamation, "SuperQAT"
End Sub
