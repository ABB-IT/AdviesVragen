Imports Arco.ABB.Common

<Serializable()> _
Public Class AlgemeenScript_OnEntry
    Inherits AdviesVragenEventHandler

    Public Overrides Sub ExecuteCode(WFCurrentCase As Doma.Library.Routing.cCase)
        Dim loDeadLineScript As ZetDeadline = New ZetDeadline
        loDeadLineScript.Execute(WFCurrentCase)

        '  Toewijzigingen.CascadeToewijzing(WFCurrentCase)
        WFCurrentCase.SetProperty("Ik keur de antwoordbrief goed", "neen")
        WFCurrentCase.SetProperty("laatste goedkeurder?", "neen")
        WFCurrentCase.SetProperty("Stuur dossier naar", "Nee")

        If WFCurrentCase.GetProperty(Of Boolean)("GMgezien") Then
            WFCurrentCase.SetProperty("Naar Gouverneur/Minister", True)
        End If

        'WFCurrentCase.RecalculateWork() try to avoid!!
    End Sub

    Public Overrides ReadOnly Property Name As String
        Get
            Return "AlgemeenScript_OnEntry"
        End Get
    End Property
End Class

