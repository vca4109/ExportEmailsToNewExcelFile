Sub ExportEmailsToNewExcelFile()
    Dim olApp As Outlook.Application
    Dim olNamespace As Outlook.NameSpace
    Dim olFolder As Outlook.MAPIFolder
    Dim olMail As Object ' Handle different types of items
    Dim xlApp As Object
    Dim xlBook As Object
    Dim xlSheet As Object
    Dim i As Integer
    Dim FilePath As String
    Dim ItemCount As Integer
    
    ' Path for the new Excel file in the Downloads folder
    FilePath = "D:\DATA\U_ANVI\Downloads\EmailSummary.xlsx"
    
    ' Initialize Outlook objects
    Set olApp = Outlook.Application
    Set olNamespace = olApp.GetNamespace("MAPI")
    Set olFolder = olNamespace.GetDefaultFolder(olFolderInbox) ' Inbox folder
    
    ' Initialize Excel objects
    Set xlApp = CreateObject("Excel.Application")
    
    ' Create a new Excel workbook
    Set xlBook = xlApp.Workbooks.Add
    Set xlSheet = xlBook.Sheets(1)
    
    ' Set headers in the Excel sheet
    xlSheet.Cells(1, 1).Value = "Date"
    xlSheet.Cells(1, 2).Value = "From"
    xlSheet.Cells(1, 3).Value = "To"
    xlSheet.Cells(1, 4).Value = "Subject"
    xlSheet.Cells(1, 5).Value = "Summary"
    
    ' Count items in the folder
    ItemCount = olFolder.Items.Count
    
    ' Debug: Show the number of items in the folder
    Debug.Print "Number of items in folder: " & ItemCount
    
    ' Loop through emails in the Inbox
    i = 2 ' Start writing from row 2 to leave row 1 for headers
    For Each olMail In olFolder.Items
        ' Check if the item is a mail item
        If TypeOf olMail Is Outlook.MailItem Then
            ' Debug: Show email subject in Immediate Window
            Debug.Print "Processing email: " & olMail.Subject
            
            ' Write email details to Excel
            xlSheet.Cells(i, 1).Value = olMail.ReceivedTime ' Date
            xlSheet.Cells(i, 2).Value = olMail.SenderName ' From
            xlSheet.Cells(i, 3).Value = olMail.To ' To
            xlSheet.Cells(i, 4).Value = olMail.Subject ' Subject
            
            ' Basic summarization: Get the first 300 characters of the email body
            xlSheet.Cells(i, 5).Value = Left(olMail.Body, 300) ' Summary
            
            ' Increment row counter
            i = i + 1
        Else
            ' Debug: Show non-mail item type
            Debug.Print "Non-mail item encountered"
        End If
    Next olMail
    
    ' Save and close the Excel workbook
    xlBook.SaveAs FilePath
    xlBook.Close
    xlApp.Quit
    
    ' Clean up
    Set olApp = Nothing
    Set olNamespace = Nothing
    Set olFolder = Nothing
    Set xlApp = Nothing
    Set xlBook = Nothing
    Set xlSheet = Nothing
    
    ' Notify user
    MsgBox "Emails have been exported to: " & FilePath
End Sub

