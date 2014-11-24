Imports Arco.Doma.Library.baseObjects
Imports Arco.Doma.Library.Routing
Imports Arco.ABB.Common
Imports Arco.Doma.Library

<Serializable()> _
Public Class AlgemeenScript_OnKeep
    Inherits AdviesVragenEventHandler

    Public Overrides Sub ExecuteCode(WFCurrentCase As Doma.Library.Routing.cCase)
        OphalenGegevens(WFCurrentCase)

        SetWeergaveInfoVeld(WFCurrentCase)
        Toewijzigingen.CascadeToewijzing(WFCurrentCase)
    End Sub

    Sub OphalenGegevens(ByRef WFCurrentCase As Doma.Library.Routing.cCase)
        Dim lsvraagsteller, lscontactpersoon As String

        ' ophalen gegevens vraagsteller
        lsvraagsteller = WFCurrentCase.GetProperty("Zoek vraagsteller").ToString
        If Not String.IsNullOrEmpty(lsvraagsteller) Then
            Dim loContact As ContactPersoon = ContactPersoon.GetContactPersoon(lsvraagsteller)
            If Not String.IsNullOrEmpty(loContact.Naam) Then
                WFCurrentCase.SetProperty("vraagsteller_naam", loContact.Naam)
                WFCurrentCase.SetProperty("vraagsteller_voornaam", loContact.VoorNaam)
                WFCurrentCase.SetProperty("vraagsteller_straatnr", loContact.StraatNr)
                WFCurrentCase.SetProperty("vraagsteller_postnummer", loContact.PostCode)
                WFCurrentCase.SetProperty("vraagsteller_gemeente", loContact.Gemeente)
                WFCurrentCase.SetProperty("adviesvrager_telefoon", loContact.Telefoon)
                WFCurrentCase.SetProperty("vraagsteller_email", loContact.Email)
                WFCurrentCase.SetProperty("adviesvrager_fax", loContact.Fax)
            End If

            WFCurrentCase.SetProperty("Zoek vraagsteller", "")
        End If


        lscontactpersoon = WFCurrentCase.GetProperty("Zoek contactpersoon").ToString
        If Not String.IsNullOrEmpty(lscontactpersoon) Then
            Dim loContact As ContactPersoon = ContactPersoon.GetContactPersoon(lscontactpersoon)
            If Not String.IsNullOrEmpty(loContact.Naam) Then
                WFCurrentCase.SetProperty("contactpersoon_email", loContact.Email)
                WFCurrentCase.SetProperty("contactpersoon_gemeente", loContact.Gemeente)
                WFCurrentCase.SetProperty("contactpersoon_naam", loContact.Naam)
                WFCurrentCase.SetProperty("contactpersoon_postnummer", loContact.PostCode)
                WFCurrentCase.SetProperty("contactpersoon_voornaam", loContact.VoorNaam)
                WFCurrentCase.SetProperty("contactpersoon-telefoonnummer", loContact.Telefoon)
                WFCurrentCase.SetProperty("contactpersoon_straatnr", loContact.StraatNr)
            End If
            WFCurrentCase.SetProperty("Zoek contactpersoon", "")
        End If

        'ophalen bestuur : when dirty or name not filled in yet
        If WFCurrentCase.GetPropertyInfo("bestuur ID").IsDirty OrElse WFCurrentCase.GetPropertyInfo("Naam bestuur vraagsteller").isEmpty Then
            Dim lsbestuur As String = WFCurrentCase.GetProperty("bestuur ID").ToString
            If Not String.IsNullOrEmpty(lsbestuur) Then
                Dim loBestuur As Bestuur = Bestuur.GetBestuur(lsbestuur)
                If Not String.IsNullOrEmpty(loBestuur.Naam) Then
                    WFCurrentCase.SetProperty("Naam bestuur vraagsteller", loBestuur.Naam)
                    WFCurrentCase.SetProperty("bestuur_straatnr", loBestuur.StraatNr)
                    WFCurrentCase.SetProperty("bestuur_postnummer", loBestuur.PostCode)
                    WFCurrentCase.SetProperty("bestuur_gemeente", loBestuur.Gemeente)
                End If

            End If
        End If

    End Sub


    Sub SetWeergaveInfoVeld(ByRef WFCurrentCase As Doma.Library.Routing.cCase)
        WFCurrentCase.SetProperty("HTMLweergave_infoveld", Formatting.FormatAssignee(WFCurrentCase.GetProperty(Of String)("dossierbehandelaar")))
    End Sub

    Public Overrides ReadOnly Property Name As String
        Get
            Return "ALGEMEENSCRIPT ON KEEP"
        End Get
    End Property
End Class
