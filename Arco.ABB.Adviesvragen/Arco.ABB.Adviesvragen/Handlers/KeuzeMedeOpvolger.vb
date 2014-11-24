
<Serializable()> _
Public Class KeuzeMedeOpvolger
    Inherits AdviesVragenEventHandler

    Public Overrides Sub ExecuteCode(WFCurrentCase As Doma.Library.Routing.cCase)
        'todo : check if used
        Dim lsBehandelaar2 As String = WFCurrentCase.GetProperty(Of String)("lookup_dossierbehandelaar2")
        WFCurrentCase.SetProperty("hiddentoewijzing3", WFCurrentCase.GetProperty(Of String)("Dienst/TEAM/Cel2"))
        WFCurrentCase.SetProperty("hiddentoewijzing4", lsBehandelaar2)
        If Not String.IsNullOrEmpty(lsBehandelaar2) Then
            WFCurrentCase.SetProperty("dossierbehandelaar2", lsBehandelaar2)
        End If

    End Sub

    Public Overrides ReadOnly Property Name As String
        Get
            Return "Keuze medeopvolger"
        End Get
    End Property
End Class
