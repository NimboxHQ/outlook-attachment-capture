Public Sub saveAttachtoDisk (itm As Outlook.MailItem)
    Dim objAtt As Outlook.Attachment
    Dim saveFolder As String
    Dim dateFormat
    dateFormat = Format(Now, "yyyy-mm-dd H-mm")
    saveFolder = "c:\[user]\Nimbox Vault\Outlook Attachments"
    For Each objAtt In itm.Attachments
        objAtt.SaveAsFile saveFolder & "\" & dateFormat & objAtt.DisplayName
        Set objAtt = Nothing
    Next
End Sub