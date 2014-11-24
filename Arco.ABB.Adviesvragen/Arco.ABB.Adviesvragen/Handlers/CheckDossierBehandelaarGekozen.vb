Imports Arco.Doma.Library

<Serializable()> _
Public Class CheckDossierBehandelaarGekozen
    Inherits AdviesVragenEventHandler

    Public Overrides Sub ExecuteCode(WFCurrentCase As Doma.Library.Routing.cCase)
        'todo : verify if logic is correct
        Dim lslookupdossierbehandelaar As String = WFCurrentCase.GetProperty(Of String)("lookup_dossierbehandelaar")
        WFCurrentCase.SetProperty("dossierbehandelaar", lslookupdossierbehandelaar) 'todo : check also if empty?

        If String.IsNullOrEmpty(lslookupdossierbehandelaar) Then
            WFCurrentCase.SetProperty("S_dossierbehandelaar?", False)
        Else
            WFCurrentCase.SetProperty("S_dossierbehandelaar?", True)
        End If
    End Sub

    Public Overrides ReadOnly Property Name As String
        Get
            Return "dossierbehandelaar gekozen? "
        End Get
    End Property
End Class

