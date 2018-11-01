Public Class Record
    Public ID As Integer
    Public BegDoc As String
    Public BegAttach As String
    Public DocDate As String
    Public DocTime As String
    Public ParentDate As String
    Public ParentTime As String
    Public ParentChild As String
    Public DateTimeItems As List(Of DateTimeItem)

    Public EmailDomains As String
    Public Designation As String

    'Public EmailDomainItems As List(Of EmailDomainItem)
End Class
