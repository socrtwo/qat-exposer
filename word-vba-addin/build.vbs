' build.vbs — Assembles SuperQAT.dotm from QATModule.bas
' Usage: cscript build.vbs
' Prerequisites: Microsoft Word installed, VBA project access enabled.

Option Explicit

Dim fso, scriptDir, basFile, outFile
Set fso = CreateObject("Scripting.FileSystemObject")
scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)
basFile   = fso.BuildPath(scriptDir, "QATModule.bas")
outFile   = fso.BuildPath(scriptDir, "SuperQAT.dotm")

If Not fso.FileExists(basFile) Then
    WScript.Echo "ERROR: QATModule.bas not found in " & scriptDir
    WScript.Quit 1
End If

' Delete old output
If fso.FileExists(outFile) Then fso.DeleteFile outFile, True

WScript.Echo "Starting Word..."
Dim word
Set word = CreateObject("Word.Application")
word.Visible = False
word.DisplayAlerts = 0  ' wdAlertsNone

Dim doc
WScript.Echo "Creating new document..."
Set doc = word.Documents.Add()

WScript.Echo "Importing VBA module..."
On Error Resume Next
doc.VBProject.VBComponents.Import basFile
If Err.Number <> 0 Then
    WScript.Echo ""
    WScript.Echo "ERROR: Could not import VBA code. (Error " & Err.Number & ": " & Err.Description & ")"
    WScript.Echo ""
    WScript.Echo "Enable VBA project access in Word:"
    WScript.Echo "  File > Options > Trust Center > Trust Center Settings > Macro Settings"
    WScript.Echo "  Check: 'Trust access to the VBA project object model'"
    doc.Close 0
    word.Quit 0
    WScript.Quit 1
End If
On Error GoTo 0

Const wdFormatXMLTemplateMacroEnabled = 13

WScript.Echo "Saving as " & outFile & " ..."
doc.SaveAs outFile, wdFormatXMLTemplateMacroEnabled
doc.Close 0  ' wdDoNotSaveChanges

WScript.Echo ""
WScript.Echo "SUCCESS: SuperQAT.dotm created!"
WScript.Echo ""
WScript.Echo "To install:"
WScript.Echo "  1. Double-click SuperQAT.dotm, or"

Dim wsh
Set wsh = CreateObject("WScript.Shell")
Dim startupFolder
startupFolder = wsh.ExpandEnvironmentStrings("%APPDATA%") & "\Microsoft\Word\STARTUP\"
WScript.Echo "  2. Copy it to: " & startupFolder

word.Quit 0
Set doc = Nothing
Set word = Nothing
WScript.Echo "Done."
