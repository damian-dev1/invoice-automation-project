Sub ImportInvoices_FromSharedMailbox_UsingParameters()
    Dim OutlookApp As Outlook.Application
    Dim ns As Outlook.Namespace
    Dim targetFolder As Outlook.Folder
    Dim item As Object, atmt As Attachment
    Dim wb As Workbook, ws As Worksheet, tbl As ListObject
    Dim paramWs As Worksheet, errorWs As Worksheet
    Dim savePath As String, filePath As String
    Dim rowData(1 To 5) As Variant
    Dim fileNameBase As String, fileExt As String
    Dim counter As Integer, errorRow As Long
    Dim folderPath As String, folderParts() As String
    Dim daysBack As Long, maxEmails As Long, emailCount As Long
    Dim i As Integer

    ' Read parameters
    Set wb = ThisWorkbook
    Set paramWs = wb.Sheets("Macro_Parameters")
    folderPath = Trim(paramWs.Range("B2").Value)
    daysBack = CLng(paramWs.Range("B4").Value)
    maxEmails = CLng(paramWs.Range("B5").Value)

    ' Prompt for output folder
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Select Output Folder"
        If .Show <> -1 Then Exit Sub
        savePath = .SelectedItems(1) & "\"
    End With

    ' Setup error log
    On Error Resume Next
    Set errorWs = wb.Sheets("Error_Log")
    If errorWs Is Nothing Then
        Set errorWs = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
        errorWs.Name = "Error_Log"
        errorWs.Range("A1:B1").Value = Array("Item Info", "Error Message")
    End If
    On Error GoTo 0
    errorRow = errorWs.Cells(errorWs.Rows.Count, 1).End(xlUp).Row + 1

    ' Setup Excel table
    Set ws = wb.Sheets("Invoice_Emails")
    Set tbl = ws.ListObjects("InvoiceTable")

    ' Setup Outlook
    Set OutlookApp = New Outlook.Application
    Set ns = OutlookApp.GetNamespace("MAPI")
    ns.Logon

    ' Navigate to folder using full path
    folderParts = Split(folderPath, "\")
    Set targetFolder = ns.Folders(folderParts(0))
    For i = 1 To UBound(folderParts)
        Set targetFolder = targetFolder.Folders(folderParts(i))
    Next i

    ' Loop through emails
    emailCount = 0
    For Each item In targetFolder.Items
        If emailCount >= maxEmails Then Exit For
        On Error GoTo ItemError
        If item.Class = olMail Then
            If item.ReceivedTime >= Date - daysBack Then
                If item.Attachments.Count > 0 Then
                    For Each atmt In item.Attachments
                        If LCase(Right(atmt.fileName, 4)) = ".pdf" Then
                            fileNameBase = Format(item.ReceivedTime, "yyyymmdd_HHmmss") & "_" & Replace(atmt.fileName, " ", "_")
                            fileExt = ".pdf"
                            filePath = savePath & fileNameBase & fileExt
                            counter = 1
                            Do While Dir(filePath) <> ""
                                filePath = savePath & fileNameBase & "_" & counter & fileExt
                                counter = counter + 1
                            Loop

                            atmt.SaveAsFile filePath

                            rowData(1) = item.Subject
                            rowData(2) = item.SenderName
                            rowData(3) = item.ReceivedTime
                            rowData(4) = atmt.fileName
                            rowData(5) = filePath

                            tbl.ListRows.Add AlwaysInsert:=True
                            tbl.ListRows(tbl.ListRows.Count).Range.Value = rowData
                            emailCount = emailCount + 1
                        End If
                    Next atmt
                End If
            End If
        End If
        GoTo ContinueLoop

ItemError:
        errorWs.Cells(errorRow, 1).Value = "Subject: " & item.Subject
        errorWs.Cells(errorRow, 2).Value = Err.Description
        errorRow = errorRow + 1
        Resume ContinueLoop

ContinueLoop:
        Err.Clear
        On Error GoTo ItemError
    Next item

    MsgBox "Done: PDFs saved to " & savePath, vbInformation
End Sub
