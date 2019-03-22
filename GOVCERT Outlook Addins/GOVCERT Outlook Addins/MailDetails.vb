Public Class MailDetails
    Public HeaderInformations As String()
    Public From As String
    Public Subject As String
    Public NumberOfAttachments As Integer

    Public Function GetSubjectLine() As String
        Return "[" + Me.Subject + "] from: " + Me.From
    End Function

End Class
