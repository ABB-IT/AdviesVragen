Imports Arco.Doma.Library

<Serializable()> _
Public Class Goedkeuren_onKeep
    Inherits AdviesVragenEventHandler

    Public Overrides Sub ExecuteCode(WFCurrentCase As Doma.Library.Routing.cCase)
        WFCurrentCase.SetProperty("hiddentoewijzing3", WFCurrentCase.GetProperty("goedkeuring_Dienst/TEAM/Cel"))
        WFCurrentCase.SetProperty("hiddentoewijzing4", WFCurrentCase.GetProperty("keuze van de goedkeurder"))
        'If Not WFCurrentCase.GetPropertyInfo("keuze van de goedkeurder").isEmpty Then
        '    WFCurrentCase.SetProperty("goedkeurder", WFCurrentCase.GetProperty("keuze van de goedkeurder"))
        'End If

        If WFCurrentCase.GetProperty(Of String)("laatste goedkeurder?") = "ja" Then
            WFCurrentCase.SetPropertyVisible("Keuze Dossierbehandelaar", True)
            WFCurrentCase.SetPropertyVisible("goedkeuring_afdeling", True)
            WFCurrentCase.SetPropertyVisible("goedkeuring_Dienst/TEAM/Cel", True)
            WFCurrentCase.SetPropertyVisible("goedkeurder", True)
        Else
            WFCurrentCase.SetPropertyVisible("Keuze Dossierbehandelaar", False)
            WFCurrentCase.SetPropertyVisible("goedkeuring_afdeling", False)
            WFCurrentCase.SetPropertyVisible("goedkeuring_Dienst/TEAM/Cel", False)
            WFCurrentCase.SetPropertyVisible("goedkeurder", False)
        End If
    End Sub

    Public Overrides ReadOnly Property Name As String
        Get
            Return "GOEDKEUREN - ON KEEP"
        End Get
    End Property
End Class

