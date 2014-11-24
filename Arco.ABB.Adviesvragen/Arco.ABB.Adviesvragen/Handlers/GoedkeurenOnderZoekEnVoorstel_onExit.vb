
<Serializable()> _
Public Class GoedkeurenOnderZoekEnVoorstel_onExit
    Inherits AdviesVragenEventHandler

    Public Overrides Sub ExecuteCode(WFCurrentCase As Doma.Library.Routing.cCase)

        If WFCurrentCase.GetPropertyInfo("goedkeurder").isEmpty Then
            WFCurrentCase.RejectComment = "geen goedkeurder aangeduid! "
        End If

    End Sub

    Public Overrides ReadOnly Property Name As String
        Get
            Return "Goedkeuren onderzoek en voorstel - ONEXIT"
        End Get
    End Property
End Class
