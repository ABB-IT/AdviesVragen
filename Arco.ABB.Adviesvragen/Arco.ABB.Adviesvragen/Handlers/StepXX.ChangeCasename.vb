Imports Arco.Doma.Library

<Serializable()> _
Public Class StepXX
    Inherits AdviesVragenEventHandler

    Public Overrides Sub ExecuteCode(WFCurrentCase As Doma.Library.Routing.cCase)
        'todo : verify + rename to ChangeCaseName

        Dim lsDossierNaam As String = String.Format("{0} - {1} Adviesvraag {2}: {3} {4}", WFCurrentCase.CaseData.Creation_Date.Substring(0, 4), WFCurrentCase.GetProperty(Of String)("S_Dossiernummer"), WFCurrentCase.GetProperty(Of String)("type vraagsteller"), WFCurrentCase.GetProperty(Of String)("type bestuur"), WFCurrentCase.GetProperty(Of String)("Naam bestuur vraagsteller"))
        '  Arco.Utils.Logging.Log("DossierNaam:" & lsDossierNaam)
        'SDP : don't save
        WFCurrentCase.Case_Name = lsDossierNaam
        'WFCurrentCase.Save()
        'WFCurrentCase.UpdateListInfo()
    End Sub

    Public Overrides ReadOnly Property Name As String
        Get
            Return "ChangeCaseName"
        End Get
    End Property
End Class
