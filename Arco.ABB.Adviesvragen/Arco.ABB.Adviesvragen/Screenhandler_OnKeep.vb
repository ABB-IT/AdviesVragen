Imports Arco.Doma.Library.baseObjects
Imports Arco.Doma.Library.Routing
Imports Arco.ABB.Common
Imports Arco.Doma.Library

<Serializable()> _
Public Class AdviesvragenScreen
    Inherits Extensibility.CaseScreenHandler

    Public Overrides Sub onBeforeRender(ByRef roScreenItems As Arco.Doma.Library.Website.ScreenItemList, ByVal voScreenMode As Arco.Doma.Library.Website.Screen.DetailScreenDisplayMode, ByVal voCase As Arco.Doma.Library.Routing.cCase)
        Arco.ABB.Common.Logging.AddToLog(voCase, "Start AdviesvragenScreen.OnBeforeRender")

        'Toon de velden van de vraagsteller
        Dim lstypevraagsteller As String = ""
        lstypevraagsteller = voCase.GetProperty(Of String)("type vraagsteller")

        If Not String.IsNullOrWhiteSpace(lstypevraagsteller) Then

            roScreenItems.Item(roScreenItems.GetFieldIndexByIdentifier("vraagsteller_naam")).Mode = Website.ScreenItem.ItemMode.Writable
            roScreenItems.Item(roScreenItems.GetFieldIndexByIdentifier("vraagsteller_voornaam")).Mode = Website.ScreenItem.ItemMode.Writable

            If MeldingsType.GetMeldingsType(lstypevraagsteller).Type = "Bestuur" Then
                Dim liBestuurIDItem As Integer = roScreenItems.GetFieldIndexByIdentifier("bestuur_ID")
                If liBestuurIDItem > 0 Then
                    roScreenItems.Item(liBestuurIDItem).Mode = Website.ScreenItem.ItemMode.Writable
                End If
                roScreenItems.Item(roScreenItems.GetFieldIndexByIdentifier("Naam bestuur vraagsteller")).Mode = Website.ScreenItem.ItemMode.Writable
                roScreenItems.Item(roScreenItems.GetFieldIndexByIdentifier("Naam organisatie vraagsteller")).Mode = Website.ScreenItem.ItemMode.Hidden

                voCase.SetProperty("Naam organisatie vraagsteller", "")
                'Dim libestuurid As Int32 = 0
                'If IsNumeric(voCase.GetProperty("bestuur ID").ToString) Then
                '    libestuurid = CInt(voCase.GetProperty("bestuur ID"))
                '    Arco.Utils.Logging.Log("libestuurid:" & libestuurid)
                '    If libestuurid <> 0 Then
                '        Call voCase.SetProperty("Naam bestuur vraagsteller", Bestuur.GetBestuur(libestuurid).Naam)
                '    End If
                'End If
            Else

                Dim liBestuurIDItem As Integer = roScreenItems.GetFieldIndexByIdentifier("bestuur_ID")
                If liBestuurIDItem > 0 Then
                    roScreenItems.Item(liBestuurIDItem).Mode = Website.ScreenItem.ItemMode.Hidden
                End If
                roScreenItems.Item(roScreenItems.GetFieldIndexByIdentifier("Naam bestuur vraagsteller")).Mode = Website.ScreenItem.ItemMode.Hidden
            End If
        Else
            roScreenItems.Item(roScreenItems.GetFieldIndexByIdentifier("Naam bestuur vraagsteller")).Mode = Website.ScreenItem.ItemMode.Hidden
            roScreenItems.Item(roScreenItems.GetFieldIndexByIdentifier("Naam organisatie vraagsteller")).Mode = Website.ScreenItem.ItemMode.Hidden
            roScreenItems.Item(roScreenItems.GetFieldIndexByIdentifier("vraagsteller_naam")).Mode = Website.ScreenItem.ItemMode.Hidden
            roScreenItems.Item(roScreenItems.GetFieldIndexByIdentifier("vraagsteller_voornaam")).Mode = Website.ScreenItem.ItemMode.Hidden
            Dim liBestuurIDItem As Integer = roScreenItems.GetFieldIndexByIdentifier("bestuur_ID")
            If liBestuurIDItem > 0 Then
                roScreenItems.Item(liBestuurIDItem).Mode = Website.ScreenItem.ItemMode.Hidden
            End If
        End If

        'toon Kies Dossierbehandelaar 
        Try
            Dim liKnop As Integer = roScreenItems.GetFieldIndexByIdentifier("Kies Dossierbehandelaar")
            Logging.AddToLog(voCase, "liKnop = " & liKnop)
            If liKnop > 0 Then
                If voCase.GetProperty(Of String)("Wie behandelt?") <> "iemand anders" AndAlso voCase.GetProperty(Of String)("Stuur dossier naar") <> "ja" Then
                    roScreenItems.Item(liKnop).Mode = Website.ScreenItem.ItemMode.Hidden
                End If
            End If
        Catch ex As Exception
            Logging.AddToLog(voCase, "Kan event Kies Dossierbehandelaar niet vinden op het scherm")
        End Try

        Try
            Dim liKnop As Integer = roScreenItems.GetFieldIndexByIdentifier("Kies Goedkeurder")
            Dim liPropOnScreen As Integer = roScreenItems.GetFieldIndexByIdentifier("Al goedgekeurd?")

            Logging.AddToLog(voCase, "liKnop = " & liKnop)
            Logging.AddToLog(voCase, "liPropOnScreen = " & liPropOnScreen)
            Dim lsGoedgekeurd As String = voCase.GetProperty(Of String)("Al goedgekeurd?")
            If liKnop > 0 Then
                If liPropOnScreen > 0 Then      ' "Al goedgekeurd?" staat op scherm: OK
                    If voCase.GetProperty(Of String)("Al goedgekeurd?") <> "Ik keur zelf goed" Then
                        roScreenItems.Item(liKnop).Mode = Website.ScreenItem.ItemMode.ReadOnly
                    Else
                        roScreenItems.Item(liKnop).Mode = Website.ScreenItem.ItemMode.Hidden
                    End If
                Else                            ' "Al goedgekeurd?" staat niet op scherm --> dit is scherm "Goedkeuring".
                    If voCase.GetProperty(Of String)("laatste goedkeurder?") <> "neen" Then
                        roScreenItems.Item(liKnop).Mode = Website.ScreenItem.ItemMode.ReadOnly
                    Else
                        roScreenItems.Item(liKnop).Mode = Website.ScreenItem.ItemMode.Hidden
                    End If
                End If
            End If
        Catch ex As Exception
            Logging.AddToLog(voCase, "Kan event Kies Goedkeurder niet vinden op het scherm")
        End Try
        'toon Goedkeurder.

        Logging.AddToLog(voCase, "Einde AdviesvragenScreen.OnBeforeRender")
    End Sub
End Class



