Imports Arco.ABB.Common

<Serializable()> _
Public Class GoedkeurderIsIkzelf
    Inherits AdviesVragenEventHandler

    Public Overrides Sub ExecuteCode(WFCurrentCase As Doma.Library.Routing.cCase)

        ' DBe 29/04/2014:
        'Dim lsKeuzeGoedkeurder As String = WFCurrentCase.GetProperty(Of String)("keuze van de goedkeurder")
        Dim lsKeuzeGoedkeurder As String = WFCurrentCase.GetProperty(Of String)("goedkeurder")

        ' DBe 20/06/2014: verkeerde toewijzing van goedkeurder
        'Dim lsDienstTeamCel As String = WFCurrentCase.GetProperty(Of String)("Dienst/TEAM/Cel")
        Dim lsDienstTeamCel As String = WFCurrentCase.GetProperty(Of String)("goedkeuring_Dienst/TEAM/Cel")

        ''''aangepast op 20110317 goedkeuren on keep geïncluded
        WFCurrentCase.SetProperty("hiddentoewijzing3", lsDienstTeamCel)
        WFCurrentCase.SetProperty("hiddentoewijzing4", lsKeuzeGoedkeurder)
        If Not String.IsNullOrEmpty(lsKeuzeGoedkeurder) Then
            WFCurrentCase.SetProperty("goedkeurder", lsKeuzeGoedkeurder)
        End If
        ' DBe 29/04/2014:
        WFCurrentCase.SetProperty("keuze van de goedkeurder", lsKeuzeGoedkeurder)

        Dim lsAlGoedgekeurd As String = WFCurrentCase.GetProperty(Of String)("Al goedgekeurd?")

        If lsAlGoedgekeurd = "Ik keur zelf goed" Then
            WFCurrentCase.SetPropertyVisible("Keuze goedkeurder", False)
            WFCurrentCase.SetProperty("Goedkeuring_afdeling", WFCurrentCase.GetProperty("afdeling"))
            ' DBe 20/06/2014: verkeerde toewijzing van goedkeurder
            'WFCurrentCase.SetProperty("goedkeuring_Dienst/TEAM/Cel", lsDienstTeamCel)
            WFCurrentCase.SetProperty("goedkeuring_Dienst/TEAM/Cel", WFCurrentCase.GetProperty("Dienst/TEAM/Cel"))
            WFCurrentCase.SetProperty("goedkeurder", WFCurrentCase.GetProperty("lookup_dossierbehandelaar"))
            WFCurrentCase.SetPropertyVisible("Goedkeuring_afdeling", False)
            WFCurrentCase.SetPropertyVisible("goedkeuring_Dienst/TEAM/Cel", False)
        Else
            If lsAlGoedgekeurd = "Er moet nog goedgekeurd worden door:" Then
                WFCurrentCase.SetPropertyVisible("Keuze goedkeurder", True)
                WFCurrentCase.SetPropertyVisible("Goedkeuring_afdeling", True)
                WFCurrentCase.SetPropertyVisible("goedkeuring_Dienst/TEAM/Cel", True)
                WFCurrentCase.SetPropertyVisible("goedkeurder", True)
                ' DBe 20/06/2014: verkeerde toewijzing van goedkeurder
                'WFCurrentCase.SetProperty("Goedkeuring_afdeling", WFCurrentCase.GetProperty("afdeling"))
                'WFCurrentCase.SetProperty("goedkeuring_Dienst/TEAM/Cel", lsDienstTeamCel)
                'WFCurrentCase.SetProperty("goedkeuring_Dienst/TEAM/Cel", WFCurrentCase.GetProperty("Dienst/TEAM/Cel"))
            Else
                WFCurrentCase.SetPropertyVisible("Keuze goedkeurder", True)
                WFCurrentCase.SetPropertyVisible("Goedkeuring_afdeling", True)
                WFCurrentCase.SetPropertyVisible("goedkeuring_Dienst/TEAM/Cel", True)
                WFCurrentCase.SetPropertyVisible("goedkeurder", True)
            End If
        End If

        ' enkel geldig voor het transfereren van het dossier
        'todo : dit is ook het transfer script (nog om te zetten)
        If WFCurrentCase.GetProperty(Of String)("Stuur dossier naar") = "ja" Then
            WFCurrentCase.SetPropertyVisible("dossierbehandelaar", True)
            WFCurrentCase.SetPropertyVisible("Dienst/TEAM/Cel", True)
            WFCurrentCase.SetPropertyVisible("afdeling", True)
            WFCurrentCase.SetPropertyVisible("Keuze Dossierbehandelaar", True)
        Else
            WFCurrentCase.SetPropertyVisible("dossierbehandelaar", False)
            WFCurrentCase.SetPropertyVisible("Dienst/TEAM/Cel", False)
            WFCurrentCase.SetPropertyVisible("afdeling", False)
            WFCurrentCase.SetPropertyVisible("Keuze Dossierbehandelaar", False)
        End If

    End Sub

    Public Overrides ReadOnly Property Name As String
        Get
            Return "Goedkeuder= ikzelf"
        End Get
    End Property
End Class

