Imports Arco.Doma.Library
Imports Arco.ABB.Common

<Serializable()> _
Public Class GeneratieBrievenAdvies
    Inherits AdviesVragenEventHandler

    Public Overrides Sub ExecuteCode(WFCurrentCase As Doma.Library.Routing.cCase)
        Dim lsHiddenAdd As String = WFCurrentCase.GetProperty(Of String)("hiddenAdd")


        'Call WFSetProperty("Keuze template","Lege brief")

        ' Dim lsKeuzetemplate As String = WFCurrentCase.GetProperty(Of String)("Keuze template")

        If lsHiddenAdd = "Add" Then
            WFCurrentCase.SetProperty("hiddenAdd", "") 'todo : convert to virtual property

            Dim lsKeuzetemplate As String = "Management Samenvatting"
            WFCurrentCase.SetProperty("Keuze template", "Management Samenvatting")


            Dim loTemplate As IBriefTemplate = Nothing
            Select Case lsKeuzetemplate
                Case "Management Samenvatting"
                    WFCurrentCase.RemoveAllFromPackage(Constants.TemplateRoutingPackage)
                    loTemplate = New ManagementSamenVattingAdviesTemplate
                Case "Lege brief"
                    loTemplate = New LegeBriefTemplate
            End Select
            If Not loTemplate Is Nothing Then
                Dim loSettings As Arco.Doma.Library.Helpers.ArcoInfo = Arco.Doma.Library.Helpers.ArcoInfo.GetParameters
                Dim lsCreatedFile As String = loTemplate.CreateFromTemplate(WFCurrentCase, loSettings)

                Dim loRoutingFile As Arco.Doma.Library.Routing.RoutingFile = Arco.Doma.Library.Routing.RoutingFile.NewFile(WFCurrentCase.Case_ID, WFCurrentCase.GetPackageInfo(Constants.TemplateRoutingPackage).ID)
                ' Vraag van ABB: formaat van naamgeving wijzigen met JJJJMMDD en alle spaties vervangen door underscores.
                Dim dTemp As DateTime = Now
                'loRoutingFile.Title = lsKeuzetemplate
                loRoutingFile.Title = lsKeuzetemplate.Replace(" ", "_") & "_" & dTemp.ToString("dd/MM/yyyy_HH:mm:ss")

                loRoutingFile.MoveFileToTargetPath = True
                loRoutingFile.TargetBasePath = Arco.Doma.FileManager.Directory.AddSlash(loSettings.GetValue("Locations", "DocPath", ""))
                loRoutingFile.Path = lsCreatedFile

                If Not String.IsNullOrEmpty(WFCurrentCase.StepExecutor) Then
                    loRoutingFile.Author = ACL.User.GetUser(WFCurrentCase.StepExecutor).USER_DISPLAY_NAME
                End If
                loRoutingFile = loRoutingFile.Save
            End If
        End If
    End Sub

    Public Overrides ReadOnly Property Name As String
        Get
            Return "GENERATIE BRIEVEN Advies"
        End Get
    End Property
End Class

