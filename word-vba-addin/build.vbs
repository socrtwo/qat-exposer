' build.vbs — Assembles SuperQAT.dot from VBA modules
' Usage: cscript build.vbs
' Prerequisites: Microsoft Word installed, VBA project access enabled.

Option Explicit

Dim fso, scriptDir, outFile
Set fso = CreateObject("Scripting.FileSystemObject")
scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)
outFile   = fso.BuildPath(scriptDir, "SuperQAT.dot")

' List of VBA modules to import
Dim modules(3)
modules(0) = "QATData1.bas"
modules(1) = "QATData2.bas"
modules(2) = "QATData3.bas"
modules(3) = "QATMain.bas"

' Verify all modules exist
Dim i, modPath
For i = 0 To UBound(modules)
    modPath = fso.BuildPath(scriptDir, modules(i))
    If Not fso.FileExists(modPath) Then
        WScript.Echo "ERROR: " & modules(i) & " not found in " & scriptDir
        WScript.Quit 1
    End If
Next

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

WScript.Echo "Importing VBA modules..."
On Error Resume Next
For i = 0 To UBound(modules)
    modPath = fso.BuildPath(scriptDir, modules(i))
    doc.VBProject.VBComponents.Import modPath
    If Err.Number <> 0 Then
        WScript.Echo "ERROR importing " & modules(i) & ": " & Err.Description
        WScript.Echo ""
        WScript.Echo "Enable VBA project access in Word:"
        WScript.Echo "  File > Options > Trust Center > Trust Center Settings > Macro Settings"
        WScript.Echo "  Check: 'Trust access to the VBA project object model'"
        doc.Close 0
        word.Quit 0
        WScript.Quit 1
    End If
    WScript.Echo "  Imported " & modules(i)
Next
On Error GoTo 0

' Save as Word 97-2003 Template (.dot) — format 1 = wdFormatTemplate
' This avoids the OOXML content-type bug with .dotm files
Const wdFormatTemplate = 1

WScript.Echo "Saving as " & outFile & " ..."
doc.SaveAs outFile, wdFormatTemplate
doc.Close 0

WScript.Echo ""
WScript.Echo "SUCCESS: SuperQAT.dot created!"
WScript.Echo ""
WScript.Echo "To install:"
WScript.Echo "  1. Double-click SuperQAT.dot, or"

Dim wsh
Set wsh = CreateObject("WScript.Shell")
Dim startupFolder
startupFolder = wsh.ExpandEnvironmentStrings("%APPDATA%") & "\Microsoft\Word\STARTUP\"
WScript.Echo "  2. Copy it to: " & startupFolder

word.Quit 0
Set doc = Nothing
Set word = Nothing
WScript.Echo "Done."
