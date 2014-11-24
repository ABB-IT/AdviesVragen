Imports Arco.Doma.Library

<Serializable()> _
Public Class Goedkeuren_OnExit
    Inherits AdviesVragenEventHandler


    Public Overrides Sub ExecuteCode(WFCurrentCase As Doma.Library.Routing.cCase)

        If WFCurrentCase.GetProperty(Of String)("laatste goedkeurder?") = "ja" AndAlso WFCurrentCase.GetPropertyInfo("keuze van de goedkeurder").isEmpty Then
            WFCurrentCase.RejectComment = "Geen goedkeurder aangeduid! "
        End If
    End Sub

    Public Overrides ReadOnly Property Name As String
        Get
            Return "Goedkeuren_OnExit"
        End Get
    End Property
End Class

