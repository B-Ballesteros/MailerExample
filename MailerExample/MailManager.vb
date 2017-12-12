Imports Microsoft.Office.Interop.Outlook

Public Class MailManager
    ''' <summary>
    ''' This function opens the outlook application or gets the instance of outlook application and sends an email based on the object
    ''' </summary>
    ''' <param name="eMail">Object that represents an email</param>
    Public Shared Sub Send(ByVal eMail As eMail)
        'Open outlook 
        Dim app As _Application
        app = New Application
        'Create a new email object
        Dim mailItem As _MailItem
        mailItem = app.CreateItem(OlItemType.olMailItem)

        'Mail filling
        For Each recipient As String In eMail.Recipients 'loop through all recipients
            mailItem.Recipients.Add(recipient)
        Next
        'email subject
        mailItem.Subject = eMail.Subject
        'email body
        mailItem.Body = eMail.Body

        'Attachment
        If eMail.AttachmentPath <> "" Then 'check if has a attahment and adds it
            Dim bodyLength = eMail.AttachmentPath.Length
            Dim attachments = mailItem.Attachments
            Dim attachment As Attachment
            attachment = attachments.Add(eMail.AttachmentPath, , bodyLength, eMail.DisplayName)
        End If
        'actual mail sending
        mailItem.Send()
    End Sub
End Class
