Imports Arco.Doma.Library
Imports Arco.ABB.Common

<Serializable()> _
Public Class WieBehandelt
    Inherits AdviesVragenEventHandler

    Public Overrides Sub ExecuteCode(WFCurrentCase As Doma.Library.Routing.cCase)
        'todo : verify Wie Behandelt of () Wie Behandelt
        If WFCurrentCase.GetPropertyInfo("Wie behandelt?").IsDirty Then
            If WFCurrentCase.GetProperty(Of String)("Wie behandelt?") = "iemand anders" Then
                WFCurrentCase.SetProperty("S_dossierbehandelaar?", False)
                WFCurrentCase.SetPropertyVisible("dossierbehandelaar", True)
                WFCurrentCase.SetPropertyVisible("Dienst/TEAM/Cel", True)
                WFCurrentCase.SetPropertyVisible("afdeling", True)
                WFCurrentCase.SetPropertyVisible("Keuze Dossierbehandelaar", True)
                Toewijzigingen.CascadeToewijzing(WFCurrentCase)
            Else
                WFCurrentCase.SetPropertyVisible("Keuze Dossierbehandelaar", False)
                WFCurrentCase.SetPropertyVisible("dossierbehandelaar", False)
                WFCurrentCase.SetPropertyVisible("Dienst/TEAM/Cel", False)
                WFCurrentCase.SetPropertyVisible("afdeling", False)
                WFCurrentCase.SetProperty("dossierbehandelaar", WFCurrentCase.CaseData.Created_By)
                WFCurrentCase.SetProperty("lookup_dossierbehandelaar", WFCurrentCase.CaseData.Created_By)

                Dim loTeam As ACL.RoleMemberList.RoleMemberInfo = Rollen.GetDienstTeam(WFCurrentCase.CaseData.Created_By)
                If loTeam.ROLE_ID > 0 Then
                    WFCurrentCase.SetProperty("Dienst/TEAM/Cel", "(Role) " & loTeam.ROLENAME)
                    Dim lsAfdeling As String = Rollen.GetAfdelingFromRol(loTeam.ROLE_ID)
                    If Not String.IsNullOrEmpty(lsAfdeling) Then
                        WFCurrentCase.SetProperty("afdeling", "(Role) " & lsAfdeling)
                    End If
                End If

                WFCurrentCase.SetProperty("S_dossierbehandelaar?", True)
                WFCurrentCase.SetPropertyVisible("dossierbehandelaar", True)

            End If
        End If

    End Sub

    Public Overrides ReadOnly Property Name As String
        Get
            Return "Wie behandelt?"
        End Get
    End Property
End Class
