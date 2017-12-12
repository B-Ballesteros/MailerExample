Public Class Form1
    Private Sub sendEmail()
        'Create an email object for test porpuse
        Dim eMail As New eMail
        With eMail
            .Recipients = New List(Of String)
            .Recipients.Add("mail@example.com") 'Make sure the address or addresses exists
            .Body = "FYI" 'email body
            .Subject = "Test" 'email subject
            .AttachmentPath = "C:\Test\people.xml" 'full path of attachment location
            .DisplayName = "People" 'attachment display name on the e-mail
        End With
        'This function sends the eMail
        MailManager.Send(eMail)
    End Sub

    Private Sub SendBtn_Click(sender As Object, e As EventArgs) Handles SendBtn.Click
        sendEmail()
    End Sub
End Class
