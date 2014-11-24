Imports Arco.Doma.Library

<Serializable()> _
Public Class ToonWelkeMailbox
    Inherits AdviesVragenEventHandler


    Public Overrides Sub ExecuteCode(WFCurrentCase As Doma.Library.Routing.cCase)
        If WFCurrentCase.GetProperty(Of String)("Medium adviesvraag") = "mail" Then
            WFCurrentCase.SetPropertyVisible("via_welke_mailbox", True)
        Else
            WFCurrentCase.SetPropertyVisible("via_welke_mailbox", False)
            If Not WFCurrentCase.GetPropertyInfo("Medium adviesvraag").isEmpty Then
                WFCurrentCase.SetProperty("via_welke_mailbox", "")
            End If
        End If
    End Sub

    Public Overrides ReadOnly Property Name As String
        Get
            Return "Toon welke mailbox"
        End Get
    End Property
End Class
