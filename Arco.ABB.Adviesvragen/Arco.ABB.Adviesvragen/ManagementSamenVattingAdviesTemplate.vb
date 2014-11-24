Imports Arco.Doma.Library
Imports Arco.ABB.Common

<Serializable()> _
Public Class ManagementSamenVattingAdviesTemplate
    Inherits RTFBriefTemplate

    Public Overrides ReadOnly Property TemplateFile As String
        Get
            Return "MGTsamenvatting.rtf"
        End Get
    End Property

    Public Overrides Function ReplaceTags(ByVal vsContent As String, voCase As Doma.Library.Routing.cCase) As String

        Dim lsNaamBestuur As String
        Dim lsdatumkabinet As String
        Dim lsJaar As String


        lsdatumkabinet = voCase.GetProperty(Of String)("Datum dossier naar G/M")
        lsJaar = DateTime.Parse(voCase.CaseData.Creation_Date).Year.ToString

        If Not String.IsNullOrEmpty(lsdatumkabinet) Then
            Dim dTemp As DateTime
            If DateTime.TryParse(lsdatumkabinet, dTemp) Then
                lsdatumkabinet = dTemp.ToString("dd/MM/yyyy")
                'lsdatumkabinet = LeadingZero(Day(CDate(lsdatumkabinet)), 2) & "/" & LeadingZero(Month(CDate(lsdatumkabinet)), 2) & "/" & Year(CDate(lsdatumkabinet))
            End If
        End If

        Dim lsBehandelaar As String = voCase.GetProperty(Of String)("dossierbehandelaar")

        Dim loBehandelaar As ACL.User = ACL.User.GetUser(lsBehandelaar)
        If Not String.IsNullOrEmpty(lsBehandelaar) Then
            loBehandelaar = ACL.User.GetUser(lsBehandelaar)
        Else
            loBehandelaar = ACL.User.NewUser("")
        End If
        Dim lsNaamBehandelaar As String = loBehandelaar.USER_DISPLAY_NAME
        Dim lsVoorEnAchterNaamBehandelaar As String = loBehandelaar.USER_FIRSTNAME & " " & loBehandelaar.USER_LASTNAME
        If Not String.IsNullOrWhiteSpace(lsVoorEnAchterNaamBehandelaar) Then
            lsNaamBehandelaar = lsVoorEnAchterNaamBehandelaar
        End If

        'start replaceing tags

        '#postdatum klacht#
        '#Naam DBH#
        '#telefoonnummer DBH#
        '#Mailadres DBH#
        '#Betreft#
        '#jaar# 
        '#S_Dossiernummer#
        '#klager_voornaam#
        '#klager_naam#
        '#klager_straatnr#
        '#klager_postnummer#
        '#klager_woonplaats#
        '#BB_AFDELING(NAAM)#
        '#BB_AFDELING(STRAATNR)#
        '#BB_AFDELING(POSTCODE)# 
        '#BB_AFDELING(GEMEENTE)#
        '#BB_AFDELING(TEL)#
        '#BB_AFDELING(FAX)#
        '#BB_AFDELING(EMAIL)#
        '#BB_AFDELING(GOUV/MINIST)#	
        '#trefwoorden#
        'trefwoorden ophalen 23/03/2011

        Dim lstrefwoorden As String = Trefwoorden.GetTrefwoordenLijst(voCase)

        'lsbinnenkomstdatum = vocaseGetProperty("Datum binnenkomst adviesvraag")
        '	If Isdate(lsbinnenkomstdatum) Then
        '		lsbinnenkomstdatum = DateAdd("d",30,lsbinnenkomstdatum) 
        '		lsbinnenkomstdatum = Day(CDate(lsbinnenkomstdatum)) & "/"& Month(CDate(lsbinnenkomstdatum)) & "/"& Year(CDate(lsbinnenkomstdatum))	

        '	End If
        'todo : verify if can be empty + leading 0
        Dim lsStepDueDate As String
        Try
            lsStepDueDate = DateTime.Parse(voCase.Step_DueDate).Date.ToString("dd/MM/yyyy")
        Catch ex As Exception
            lsStepDueDate = ""
        End Try

        vsContent = vsContent.Replace("#Type bestuur#", voCase.GetProperty(Of String)("type vraagsteller"))

        lsNaamBestuur = voCase.GetProperty(Of String)("Naam bestuur vraagsteller")
        If Not String.IsNullOrEmpty(lsNaamBestuur) Then
            vsContent = vsContent.Replace("#Bestuur-naam#", lsNaamBestuur)
        Else
            vsContent = vsContent.Replace("#Bestuur-naam#", voCase.GetProperty(Of String)("Naam organisatie vraagsteller"))
        End If
        vsContent = vsContent.Replace("#aard_vraag#", voCase.GetProperty(Of String)("aard_vraag_lijst"))


        vsContent = vsContent.Replace("#datum binnenkomst#", lsStepDueDate)

        vsContent = vsContent.Replace("#trefwoorden#", lstrefwoorden)
        vsContent = vsContent.Replace("#datum kabinet#", lsdatumkabinet)
        vsContent = vsContent.Replace("#Naam DBH#", lsNaamBehandelaar)

        vsContent = vsContent.Replace("#telefoonnummer DBH#", loBehandelaar.USER_PHONE)
        vsContent = vsContent.Replace("#Mailadres DBH#", loBehandelaar.USER_MAIL)
        vsContent = vsContent.Replace("#Betreft#", voCase.GetProperty(Of String)("betreft"))
        vsContent = vsContent.Replace("#jaar# ", lsJaar)
        vsContent = vsContent.Replace("#S_Dossiernummer#", voCase.GetProperty(Of String)("S_Dossiernummer"))
        vsContent = vsContent.Replace("#systeemdatum#", System.DateTime.Now.ToString("dd/MM/yyyy"))


        Dim loAfdeling As Afdeling = Afdeling.GetAfdeling(voCase.GetProperty(Of String)("afdeling"))

        vsContent = vsContent.Replace("#BB_AFDELING(NAAM)#", loAfdeling.Naam)
        vsContent = vsContent.Replace("#BB_AFDELING(STRAATNR)#", loAfdeling.StraatNr)
        vsContent = vsContent.Replace("#BB_AFDELING(GEMEENTE)#", loAfdeling.Gemeente)
        vsContent = vsContent.Replace("#BB_AFDELING(POSTCODE)#", loAfdeling.PostCode)
        vsContent = vsContent.Replace("#BB_AFDELING(TEL)#", loAfdeling.Telefoon)
        vsContent = vsContent.Replace("#BB_AFDELING(FAX)#", loAfdeling.Fax)
        vsContent = vsContent.Replace("#BB_AFDELING(EMAIL)#", loAfdeling.Email)

        vsContent = vsContent.Replace("#NAAM_Afdelingshoofd#", loAfdeling.NaamAfdelingshoofd)
        vsContent = vsContent.Replace("#AANSPREEKTITEL#", loAfdeling.AanspreekTitel)

        Dim lsSamenvatting As String
        Dim lsKennisdeling2 As String
        Dim lsKennisdeling1 As String
        Dim lsKennisdeling3 As String
        Dim lsKennisdeling4 As String
        lsSamenvatting = voCase.GetProperty(Of String)("Samenvatting")

        If voCase.GetProperty(Of Boolean)("lbKennisdeling2") = False Then
            lsKennisdeling2 = "Neen"
        Else
            lsKennisdeling2 = "    Ja"
        End If

        If voCase.GetProperty(Of Boolean)("lbKennisdeling1") = False Then
            lsKennisdeling1 = "Neen"
        Else
            lsKennisdeling1 = "    Ja"
        End If

        If voCase.GetProperty(Of Boolean)("lbKennisdeling3") = False Then
            lsKennisdeling3 = "Neen"
        Else
            lsKennisdeling3 = "    Ja"
        End If
        If voCase.GetProperty(Of Boolean)("lbKennisdeling4") = False Then
            lsKennisdeling4 = "Neen"
        Else
            lsKennisdeling4 = "    Ja"
        End If

        vsContent = vsContent.Replace("#Samenvatting#", lsSamenvatting)
        vsContent = vsContent.Replace("#Kennisdeling2#", lsKennisdeling2)
        vsContent = vsContent.Replace("#Kennisdeling1#", lsKennisdeling1)
        vsContent = vsContent.Replace("#Kennisdeling3#", lsKennisdeling3)
        vsContent = vsContent.Replace("#Kennisdeling4#", lsKennisdeling4)

        Return vsContent
    End Function

End Class

