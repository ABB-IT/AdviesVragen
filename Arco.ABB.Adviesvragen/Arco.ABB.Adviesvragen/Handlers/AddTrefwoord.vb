Imports Arco.ABB.Common

<Serializable()> _
Public Class AddTrefwoord
    Inherits AdviesVragenEventHandler

    Public Overrides Sub ExecuteCode(WFCurrentCase As Doma.Library.Routing.cCase)
        Dim liTrefwoordID As String = WFCurrentCase.GetProperty(Of String)("trefwoorden")
        If Not String.IsNullOrEmpty(liTrefwoordID) Then
            Dim lsTrefwoord As String = Trefwoorden.GetTrefWoord(liTrefwoordID)
            If Not String.IsNullOrEmpty(lsTrefwoord) Then
                'voeg trefwoord toe aan tabel
                Dim litableID As Integer = WFCurrentCase.GetPropertyInfo("trefwoordenlijst").PROP_ID
                Dim liRowID As Integer = 0
                WFCurrentCase.CreateRowInTable(litableID, liRowID, 0)
                WFCurrentCase.SetProperty("trefwoord", lsTrefwoord, liRowID, litableID)

                WFCurrentCase.SetProperty("trefwoorden", "")

            End If

        End If

    End Sub

    Public Overrides ReadOnly Property Name As String
        Get
            Return "AddTrefwoord"
        End Get
    End Property
End Class
