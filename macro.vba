Sub MailMergeToPdfBasic()
    ' Macro to generate individual PDFs from a Mail Merge using SaveAs2

    Dim masterDoc As Document, singleDoc As Document
    Dim lastRecordNum As Long
    Dim filePath As String
    Dim fileName As String

    ' Set the active document as master
    Set masterDoc = ActiveDocument

    ' Get the last active record number
    masterDoc.MailMerge.DataSource.ActiveRecord = wdLastRecord
    lastRecordNum = masterDoc.MailMerge.DataSource.ActiveRecord

    ' Start from the first active record
    masterDoc.MailMerge.DataSource.ActiveRecord = wdFirstRecord

    ' Loop through all records
    Do While masterDoc.MailMerge.DataSource.ActiveRecord <= lastRecordNum
        masterDoc.MailMerge.Destination = wdSendToNewDocument
        masterDoc.MailMerge.DataSource.FirstRecord = masterDoc.MailMerge.DataSource.ActiveRecord
        masterDoc.MailMerge.DataSource.LastRecord = masterDoc.MailMerge.DataSource.ActiveRecord
        masterDoc.MailMerge.Execute False

        ' Set newly created document as active
        Set singleDoc = ActiveDocument

        ' Generate file path and file name
        filePath = masterDoc.MailMerge.DataSource.DataFields("PdfFolderPath").Value
        fileName = masterDoc.MailMerge.DataSource.DataFields("PdfFileName").Value & ".pdf"

        ' Ensure directory exists
        If Dir(filePath, vbDirectory) = "" Then MkDir filePath

        ' Debug: Print Active Printer
        Debug.Print "Active Printer: " & Application.ActivePrinter
        
         Call SetImageQuality
         
         singleDoc.PrintOut PrintToFile:=True, OutputFileName:=filePath & "\" & fileName
         

        ' Save as PDF using SaveAs2
       

        ' Close the document without saving
        singleDoc.Close False

        ' Move to the next record
        If masterDoc.MailMerge.DataSource.ActiveRecord >= lastRecordNum Then
            Exit Do
        Else
            masterDoc.MailMerge.DataSource.ActiveRecord = wdNextRecord
        End If
    Loop

    MsgBox "Mail Merge PDFs successfully created!", vbInformation, "Process Complete"

End Sub


Sub SetImageQuality()
    Dim opt As Options
    Set opt = Application.Options
    
    opt.PasteFormatBetweenDocuments = wdKeepSourceFormatting ' Keeps original quality
    opt.PasteFormatBetweenStyledDocuments = wdKeepSourceFormatting
    opt.PasteFormatFromExternalSource = wdKeepSourceFormatting
End Sub
