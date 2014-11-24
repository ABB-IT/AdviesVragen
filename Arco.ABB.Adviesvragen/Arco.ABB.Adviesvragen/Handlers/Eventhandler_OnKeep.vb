' **************************************************************************
' Project naming convention : Global Onkeep for  procedure Adviesvragen
' **************************************************************************
' Author : Geoffrey
' Created by         on   
' Modified by        on 
' Description :
' **************************************************************************
Imports Arco.Doma.Library.baseObjects
Imports Arco.Doma.Library.Routing
Imports Arco.ABB.Common
Imports Arco.Doma.Library


<Serializable()> _
Public Class Eventhandler_OnKeep
    Inherits AdviesVragenEventHandler


    Public Overrides Sub ExecuteCode(WFCurrentCase As Doma.Library.Routing.cCase)
        'todo : cleanup if used

        'Toon de velden van de vraagsteller
        Dim lstypevraagsteller As String = ""
        lstypevraagsteller = WFCurrentCase.GetProperty("type vraagsteller").ToString
        If WFCurrentCase.GetProperty("type vraagsteller").ToString.Trim <> "" Then
            WFCurrentCase.SetPropertyVisible("vraagsteller_naam", True)
            WFCurrentCase.SetPropertyVisible("vraagsteller_voornaam", True)

            If MeldingsType.GetMeldingsType(lstypevraagsteller).Type = "Bestuur" Then
                '		msgbox " bestuur" & WFGetProperty("Naam bestuur vraagsteller")

                WFCurrentCase.SetPropertyVisible("bestuur ID", True)
                WFCurrentCase.SetPropertyVisible("Naam bestuur vraagsteller", True)
                WFCurrentCase.SetPropertyVisible("Naam organisatie vraagsteller", False)
                WFCurrentCase.SetProperty("Naam organisatie vraagsteller", "")
                Dim libestuurid As Int32 = 0
                If IsNumeric(WFCurrentCase.GetProperty("bestuur ID").ToString) Then
                    libestuurid = CInt(WFCurrentCase.GetProperty("bestuur ID"))
                    If libestuurid <> 0 Then
                        Call WFCurrentCase.SetProperty("Naam bestuur vraagsteller", Bestuur.GetBestuur(libestuurid.ToString()).Naam)
                    End If
                End If
            Else

                WFCurrentCase.SetPropertyVisible("bestuur ID", False)
                WFCurrentCase.SetPropertyVisible("Naam bestuur vraagsteller", False)
                WFCurrentCase.SetProperty("bestuur ID", "")
                WFCurrentCase.SetProperty("Naam bestuur vraagsteller", "")
                WFCurrentCase.SetPropertyVisible("Naam organisatie vraagsteller", True)

            End If
        Else

            WFCurrentCase.SetPropertyVisible("Naam bestuur vraagsteller", False)
            WFCurrentCase.SetPropertyVisible("Naam organisatie vraagsteller", False)
            WFCurrentCase.SetPropertyVisible("vraagsteller_naam", False)
            WFCurrentCase.SetPropertyVisible("vraagsteller_voornaam", False)
            WFCurrentCase.SetPropertyVisible("bestuur ID", False)
            WFCurrentCase.SetProperty("bestuur ID", "")
            WFCurrentCase.SetProperty("Naam bestuur vraagsteller", "")
            WFCurrentCase.SetProperty("Naam organisatie vraagsteller", "")
            WFCurrentCase.SetProperty("vraagsteller_naam", "")
            WFCurrentCase.SetProperty("vraagsteller_voornaam", "")
        End If
        ' Einde tonen vraagsteller

    End Sub

    Public Overrides ReadOnly Property Name As String
        Get
            Return "Eventhandler_OnKeep"
        End Get
    End Property
End Class
