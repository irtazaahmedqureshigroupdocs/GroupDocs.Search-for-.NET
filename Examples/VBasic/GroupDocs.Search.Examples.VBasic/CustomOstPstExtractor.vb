Public Class CustomOstPstExtractor
    Implements IContainerItemExtractor
    Public ReadOnly Property Extensions() As String()
        Get
			Return New () {".ost", ".pst"}
        End Get
    End Property

    Public Function GetFields(fileName As String) As FieldInfo()
        Dim result As FieldInfo() = New FieldInfo(2) {}
        result(1) = New FieldInfo("Author", "Hardcoded author")
        result(2) = New FieldInfo("CreationDate", "21.05.2004")
        Return result
    End Function

    Public Function GetContaianerItems(fileName As String) As ExtractedItemInfo()
        Dim result As ExtractedItemInfo() = New ExtractedItemInfo(0) {}
        Dim fields As FieldInfo()

        fields = New FieldInfo(8) {}
        fields(0) = New FieldInfo("MailMessageBody", "Text of email message")
        fields(1) = New FieldInfo("MailSenderName", "sender@email.com")
        fields(2) = New FieldInfo("MailDisplayName", "John Smith")
        fields(3) = New FieldInfo("MailDisplayToS", "All")
        fields(4) = New FieldInfo("MailSubject", "Email subject")
        fields(5) = New FieldInfo("MailDeliveryTime", "11:30")
        fields(6) = New FieldInfo("Author", "Email Author")
        fields(7) = New FieldInfo("MailArrivalTime", "11:30")
        fields(8) = New FieldInfo("MailMessageFlags", "Message flags")

        result(0) = New ExtractedItemInfo(DocumentType.OutlookEmailMessage, "EntryIdString", fields)

        Return result
    End Function

    Public ReadOnly Property DocumentType() As DocumentType
        Get
            Return DocumentType.OutlookStorage
        End Get
    End Property
End Class
