Imports Arco.Doma.Library
Imports Arco.ABB.Common

<Serializable()> _
Public Class AlgemeenScript_OnExit
    Inherits AdviesVragenEventHandler

    Public Overrides Sub ExecuteCode(WFCurrentCase As Doma.Library.Routing.cCase)
        'If WFCurrentCase.GetProperty(Of String)("Datum binnenkomst adviesvraag") > WFCurrentCase.GetProperty(Of String)("registratiedatum") Then
        '    Logging.AddToLog(WFCurrentCase, "Rejected : Check de datum van binnenkomst advies. Deze moet kleiner of gelijk zijn aan de registratiedatum ")
        '    Logging.AddToLog(WFCurrentCase, "Datavalues: Datum binnenkomst adviesvraag = " & WFCurrentCase.GetProperty(Of String)("Datum binnenkomst adviesvraag"))
        '    Logging.AddToLog(WFCurrentCase, "Datavalues: registratiedatum = " & WFCurrentCase.GetProperty(Of String)("registratiedatum"))
        '    WFCurrentCase.RejectComment = "Check de datum van binnenkomst advies. Deze moet kleiner of gelijk zijn aan de registratiedatum "
        'End If
        If WFCurrentCase.GetProperty(Of Date)("Datum binnenkomst adviesvraag") > WFCurrentCase.GetProperty(Of Date)("registratiedatum") Then
            Logging.AddToLog(WFCurrentCase, "Rejected : Check de datum van binnenkomst advies. Deze moet kleiner of gelijk zijn aan de registratiedatum ")
            Logging.AddToLog(WFCurrentCase, "Datavalues: Datum binnenkomst adviesvraag = " & WFCurrentCase.GetProperty(Of Date)("Datum binnenkomst adviesvraag"))
            Logging.AddToLog(WFCurrentCase, "Datavalues: registratiedatum = " & WFCurrentCase.GetProperty(Of Date)("registratiedatum"))
            WFCurrentCase.RejectComment = "Check de datum van binnenkomst advies. Deze moet kleiner of gelijk zijn aan de registratiedatum "
        End If
    End Sub

    Public Overrides ReadOnly Property Name As String
        Get
            Return "ALGEMEEN SCRIPT - ONEXIT"
        End Get
    End Property
End Class

