Imports Arco.Doma.Library
Imports Arco.ABB.Common

<Serializable()> _
Public Class ZetDeadline
    Inherits AdviesVragenEventHandler

    Public Overrides Sub ExecuteCode(WFCurrentCase As Doma.Library.Routing.cCase)
        'Deadline zetten 
        '30 dagen na binnenkomst    
        Dim gdDatumBinnenkomstAfdeling As String = WFCurrentCase.GetProperty(Of String)("Datum binnenkomst adviesvraag")
        If WFCurrentCase.GetProperty(Of Boolean)("Complexiteit") Then
            WFCurrentCase.Step_DueDate = ""
        Else
            If Not String.IsNullOrEmpty(gdDatumBinnenkomstAfdeling) Then
                Dim datumBinnenkomst As System.DateTime
                If DateTime.TryParse(gdDatumBinnenkomstAfdeling, datumBinnenkomst) Then
                    WFCurrentCase.Step_DueDate = datumBinnenkomst.AddDays(30).ToString("yyyy-MM-dd") & " 23:59:59"
                Else
                    WFCurrentCase.Step_DueDate = ""
                End If
            Else
                WFCurrentCase.Step_DueDate = ""
            End If
        End If
        ' Wijziging DBe 26/06/2014: eerst een deadline tonen (zie originele code hierboven, daarna de tekst "geen deadline" tonen, dan opnieuw de code hierboven tonen.
        'Select Case WFCurrentCase.CurrentStep.Step_Name.ToUpper
        '    Case "AFSLUITEN DOSSIER", "STOPZETTEN DOSSIER"
        '        WFCurrentCase.Step_DueDate = ""
        'End Select

    End Sub

    Public Overrides ReadOnly Property Name As String
        Get
            Return "DEADLINE ZETTEN"
        End Get
    End Property
End Class

