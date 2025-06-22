# Outlook Invoice Importer VBA Macro

This VBA macro automates the process of extracting PDF invoice attachments from a shared Outlook mailbox and logging them into an Excel table. It uses a control sheet to dynamically configure parameters such as the target mail folder, output directory, date range, and email limit.

---

## Project Structure

- **Macro_Parameters**: Sheet to configure runtime parameters.
- **Invoice_Emails**: Sheet with a named table `InvoiceTable` to log extracted data.
- **Error_Log**: Sheet to capture any errors encountered during processing.

---

## Setup Instructions

1. **Create the Control Sheet**

   Add a worksheet named `Macro_Parameters` with the following layout:

   | Parameter                | Value                                |
   |--------------------------|--------------------------------------|
   | Mail Folder to Scan      | invoices.dropship\Inbox\Printed    |
   | Output Folder            | *(Leave blank – will use dialog)*    |
   | Number of Days           | 2                                    |
   | Maximum Number of Emails | 100                                  |

2. **Create the Output Table**

   Add a worksheet named `Invoice_Emails` and insert a table named `InvoiceTable` with 5 columns:

   - Subject
   - Sender
   - ReceivedTime
   - FileName
   - FilePath

3. **Add the Macro**

   Open the VBA editor (`Alt + F11`), insert a new module, and paste the macro code.

---

## ▶️ How to Use

1. Fill in the `Macro_Parameters` sheet.
2. Run the macro `ImportInvoices_FromSharedMailbox_UsingParameters`.
3. Select the output folder when prompted.
4. The macro will:
   - Connect to the specified Outlook folder.
   - Filter emails by date and attachment type.
   - Save PDF attachments to the selected folder.
   - Log metadata into the `InvoiceTable`.
   - Log any errors to the `Error_Log` sheet.

---

## Customization

- **Folder Path**: Use full path format like `MailboxName\Inbox\Subfolder`.
- **Attachment Filter**: Modify the file extension check to support other formats.
- **Logging**: Extend the `Error_Log` sheet to include timestamps or categories.

---

## Notes

- Performance may degrade with large mailboxes.
- Ensure Outlook is open and connected.
- The macro uses `Application.FileDialog` for folder selection.

---

## 📞 Support

For issues or enhancements, please contact Damian Damjanovic.

```vb
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
```
