' Visual Basic Script for automatically saving screenshots taken by the Window's Snipping Tool
' v0.1
'
' By Matia Choi
' Jan 8, 2024
' 
'

' Define the source and destination directories
Const SourceDir = "C:\Users\matia\AppData\Local\Microsoft\Windows\ActionCenterCache"
Const DestDir = "C:\Users\matia\Pictures\Screenshots"

' Create the FileSystemObject
Set fso = CreateObject("Scripting.FileSystemObject")

Set WshShell = WScript.CreateObject("WScript.Shell")

' Show a startup dialog
MsgBox  "Screenshot Copier by Matia Choi" & vbCrLf & "The script is starting and will monitor: " & SourceDir & "And will copy the images to: " & DestDir

' Create a new variable named lastFile and assign it an empty value
Dim lastFile
lastFile = ""

' Main loop
Do While True
    ' Check for files in the source directory
    Set folder = fso.GetFolder(SourceDir)
    Set files = folder.Files

    For Each file In files
        If Left(file.Name, 22) = "microsoft-screensketch" And file.Name <> lastFile Then

            ' Assigns the current file as the lastFile
            lastFile = file.Name

            ' File name starts with "microsoft-screensketch_", copy it
            file.Copy DestDir & "\" & file.Name

        End If
    Next

    ' Wait for a bit before checking again
    WScript.Sleep 100 ' 10 seconds
Loop