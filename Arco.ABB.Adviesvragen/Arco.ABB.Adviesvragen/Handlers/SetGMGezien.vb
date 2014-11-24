
<Serializable()> _
Public Class SetGMGezien
    Inherits AdviesVragenEventHandler

    Public Overrides Sub ExecuteCode(WFCurrentCase As Doma.Library.Routing.cCase)
        WFCurrentCase.SetProperty("GMgezien", True)
    End Sub

    Public Overrides ReadOnly Property Name As String
        Get
            Return "SetGMGezien"
        End Get
    End Property
End Class

