Sub printCertainAttachments()
    Dim toPrint As MAPIFolder
    Dim Item As MailItem
    Dim Atmt As Attachment
    Dim extension As String
    Dim FileName As String
    Dim FullFileName As String
    Dim i As Integer
    Dim j As Integer
    Dim pdfCount As Integer
    Dim wordCount As Integer
    Dim excelCount As Integer
    Dim txtCount As Integer
    Dim msgCount As Integer
    Dim total As Integer
    Dim other As Integer
    Dim username As String
    
    pdfCount = 0
    wordCount = 0
    excelCount = 0
    txtCount = 0
    msgCount = 0
    other = 0
    j = 0
    username = (Environ$("Username"))
    
    'Set folder one sublevel below Inbox
    Set toPrint = GetNamespace("MAPI").GetDefaultFolder(olFolderInbox).Folders("TO PRINT")
    'Set folder one sublevel below TO PRINT
    Set Printed = GetNamespace("MAPI").GetDefaultFolder(olFolderInbox).Folders("TO PRINT").Folders("PRINTED")
    
    'Before we start the loop, need to clear the contents of the log file from the previous print job
    clearLogFile username
    
    'Outer loop that iterates through every email in the TO PRINT folder has to go in reverse since we're moving an email at the end of every loop to the "PRINTED" folder, _
    which reduces the size of the Items array by 1 each time. Couldn't use For Each for this purpose, hence the below reverse loop.
    For i = toPrint.Items.Count To 1 Step -1
        Set Item = toPrint.Items(i)
        'Nested loop iterates through all of the attachments in a single email
        For Each Atmt In Item.Attachments
        
            FileName = Atmt.FileName
            FullFileName = "C:\Users\" & username & "\Desktop\printAttachmentsMacro\printMacro\" & j & FileName
            splitArray = Split(FileName, ".")
            extension = LCase(splitArray(UBound(splitArray)))
            'MsgBox (FullFileName)
            'If there is a duplicate file name already in ..\printMacro\ then delete it
            'checkDuplicateDelete FullFileName
            
                'If file is a PDF, Word doc, or Excel file, print it
                If extension = "pdf" Then
                    pdfCount = pdfCount + 1
                    Atmt.SaveAsFile FullFileName
                    pdftoprint = Shell("C:\Users\" & username & "\Desktop\printAttachmentsMacro\PDFtoPrinter.exe " & Chr(34) & FullFileName & Chr(34) & "")
                    'The below code is to use Adobe Reader to print PDFs. Was buggy last time used. Probably just stick with PDFtoPrinter.
                    'adobe = Shell("""C:\Program Files (x86)\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe"" /h /p """ & Chr(34) & FullFileName & Chr(34) & """", vbHide)
                ElseIf extension = "doc" Or extension = "docx" Then
                    wordCount = wordCount + 1
                    Atmt.SaveAsFile FullFileName
                    word = Shell("""C:\Program Files (x86)\Microsoft Office\root\Office16\WinWord.exe"" /q /n /mFilePrintDefault /mFileCloseOrExit """ & Chr(34) & FullFileName & Chr(34) & """", vbHide)
                ElseIf extension = "xls" Or extension = "xlsx" Then
                    excelCount = excelCount + 1
                    Atmt.SaveAsFile FullFileName
                    excel = Shell("""C:\Program Files (x86)\Microsoft Office\root\Office16\EXCEL.exe"" /q /n /mFilePrintDefault /mFileCloseOrExit " & Chr(34) & FullFileName & Chr(34) & "", vbHide)
                'ElseIf extension = "txt" Then
                    'Shouldn't need to print .txt files, but if we do, the code is ready to go
                    'txtCount = txtCount + 1
                    'Atmt.SaveAsFile FullFileName
                    'notepad = Shell("NOTEPAD /P """ & Chr(34) & FullFileName & Chr(34) & """", vbHide)
                'ElseIf extension = "msg" Then
                    'Shouldn't need to print .msg files, but if we do, the code is ready to go
                    'msgCount = msgCount + 1
                    'Atmt.SaveAsFile FullFileName
                    'mail = Shell("""C:\Program Files (x86)\Microsoft Office\root\Office16\OUTLOOK.exe"" /p """ & Chr(34) & FullFileName & Chr(34) & """", vbHide)
                Else
                    other = other + 1
                    'Send log entry to C:\Users\'username'\Desktop\printAttachmentsMacro\lastPrintMacro.log
                    logNoPrint "Could not print attachment: " & Chr(34) & FullFileName & Chr(34) & " from email: " & Chr(34) & Item & Chr(34), username
                End If
            
            j = j + 1
            
        Next
    Item.Move Printed
    Next
    
    Set toPrint = Nothing
    total = pdfCount + wordCount + excelCount
    MsgBox ("Total Printed: " & total & vbNewLine & vbNewLine & "PDFs Printed: " & pdfCount & vbNewLine & "Word Docs Printed: " & wordCount & vbNewLine & "Excel Spreadsheets Printed: " & excelCount & vbNewLine & vbNewLine & "Files Not Printed: " & other)
    Pause 5
    'delete the temp attachments in C:\Users\'username'\Desktop\printAttachmentsMacro\printMacro
    deleteFiles username
End Sub
