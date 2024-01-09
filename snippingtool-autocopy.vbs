' Visual Basic Script for automatically saving screenshots taken by the Window's Snipping Tool
' v0.1
'
' By Matia Choi
' Jan 8, 2024
' 
'

' Define the source and destination directories
Const SourceDir = "C:\Users\matia\AppData\Local\Packages\MicrosoftWindows.Client.CBS_cw5n1h2txyewy\TempState\ScreenClip"
Const DestDir = "C:\Users\matia\Pictures\Screenshots"

' Create the FileSystemObject
Set fso = CreateObject("Scripting.FileSystemObject")

' Show a startup dialog
MsgBox  "Screenshot Copier by Matia Choi" & vbCrLf & "The script is starting and will monitor: " & SourceDir & "And will copy the images to: " & DestDir

' Init date and time variable
Dim currentDateTime
currentDateTime = Now
currentDateTime = Replace(currentDateTime, "/", "-")
currentDateTime = Replace(currentDateTime, ":", "")

' Create an array to store two files since taking a screenshot generates two files.
' One high res and one low res
Dim imgArr(1)

' Main loop
Do While True
    ' Check for files in the source directory
    Set folder = fso.GetFolder(SourceDir)
    Set files = folder.Files

    ' Set the array to null
    Set imgArr(0) = Nothing
    Set imgArr(1) = Nothing

    ' Initialize img counter
    Dim imgCounter
    imgCounter = 0
    For Each file In files
        If Left(file.Name, 1) = "{" AND Right(file.Name, 3) = "png" Then
            ' Add image file to loop
            Set imgArr(img) = file
            ' Increment counter
            imgCounter = imgCounter + 1
        End If
    Next

    ' Variable to store the output file
    Dim outputFile 

    ' Check if the array is not empty
    If Not imgArr(0) Is Nothing And Not imgArr(1) Is Nothing Then

        ' Compare the sizes and assign it to outputFile accordingly
        If imgArr(0).Size > imgArr(1).Size Then
            Set outputFile = imgArr(0)
        Else
            Set outputFile = imgArr(1)
        End If

        ' Get the correct file name and copy
        currentDateTime = Now
        currentDateTime = Replace(currentDateTime, "/", "-")
        currentDateTime = Replace(currentDateTime, ":", "")
        outputFile.Copy DestDir & "\" & currentDateTime & ".png"

        ' Empty the source dir
        For Each file in files
            file.Delete(True)
        Next
    End If

    ' Wait for a bit before checking again
    WScript.Sleep 500 ' 10 seconds
Loop