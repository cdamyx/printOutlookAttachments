Sub deleteFiles(username)
'Delete all the temp attachment files in C:\Users\'username'\Desktop\printAttachmentsMacro\printMacro\
    On Error Resume Next
    Kill "C:\Users\" & username & "\Desktop\printAttachmentsMacro\printMacro\*.*"
    On Error GoTo 0
End Sub

Sub logNoPrint(logMessage, username)

    Dim LogFileName As String
    LogFileName = "C:\Users\" & username & "\Desktop\printAttachmentsMacro\lastPrintMacro.txt"
    Dim FileNum As Integer

    FileNum = FreeFile ' next file number
    Open LogFileName For Append As #FileNum ' creates the file if it doesn't exist
    Print #FileNum, logMessage ' write information at the end of the text file
    Close #FileNum ' close the file

End Sub

Sub clearLogFile(username)

    Dim LogFileName As String
    LogFileName = "C:\Users\" & username & "\Desktop\printAttachmentsMacro\lastPrintMacro.txt"
    Dim FileNum As Integer


    FileNum = FreeFile ' next file number
    Open LogFileName For Output As #FileNum
    Close #FileNum

End Sub

Sub Pause(Seconds As Single)
    Dim TimeEnd As Single
    TimeEnd = Timer + Seconds
    While Timer < TimeEnd
        DoEvents
    Wend
End Sub