''******************************************** 
''*** FUNCTIE WFCreateHistoryAction *** 
''********************************************
'Function WFCreateHistoryAction(ByVal vlActionType, ByVal vsMessage)
'    Engine.HistoryAction.ClearHistoryAction()

'    ' Case Data 
'    Engine.HistoryAction.Case_ID = WFCurrentCase.Case_ID
'    Engine.HistoryAction.Step_ID = WFCurrentCase.Step_ID
'    Engine.HistoryAction.USER_ID = "Script"
'    Engine.HistoryAction.Object_Value = WFGetStepName()
'    ' Date & Type
'    Engine.HistoryAction.ActionDate = Now()
'    Engine.HistoryAction.ActionDateEnd = Now()
'    Engine.HistoryAction.Action_Type = vlActionType
'    ' (Error)Message
'    Engine.HistoryAction.ACTION = vsMessage
'    ' Create HistoryAction
'    Engine.HistoryAction.CreateHistoryAction()
'    Engine.HistoryAction.ClearHistoryAction()
'End Function


''*****************************************
''*** FUNCTIE WFGetStepName ***
''*****************************************
'Function WFGetStepName() 'As String

'    Dim lobjSteps, mobjEngine

'    mobjEngine = Engine
'    WFGetStepName = ""
'    mobjEngine.Step.ClearStep()
'    mobjEngine.Step.Proc_ID = WFCurrentCase.Proc_ID
'    mobjEngine.Step.Step_ID = WFCurrentCase.Step_ID
'    lobjSteps = mobjEngine.Step.GetStep
'    If mobjEngine.Step.Error.Number <> 0 Then
'        'sRaiseError mobjEngine.Step.Error.Number, mobjEngine.Step.Error.Description
'        lobjSteps = Nothing
'        Exit Function
'    End If
'    If lobjSteps.Count = 1 Then
'        WFGetStepName = lobjSteps.Item(1).Step_Name
'    End If
'    mobjEngine.Step.ClearStep()
'    mobjEngine = Nothing
'End Function


'Function fsGetCaseName(ByVal vlCase_ID)

'    Dim loCases, lsCaseName
'    lsCaseName = ""

'    Engine.cCase.ClearCase()
'    Engine.cCase.Case_ID = vlCase_ID
'    loCases = Engine.cCase.Getcase(" 1 = 1")
'    If loCases.Count > 0 Then lsCaseName = loCases.Item(1).Case_Name
'    loCases = Nothing
'    fsGetCaseName = lsCaseName
'End Function


''**********************************************
''*** Function  fsDate                                 ***
''**********************************************
'Function fsDate(ByVal d)
'    d = CDate(d)
'    fsDate = Day(d) & "/" & Month(d) & "/" & Year(d)
'End Function


''**********************************
''*** FUNCTIE flGetPackID *** 
''**********************************
'Function flGetPackID(ByVal vsPackName, ByVal vlProcID)

'    Dim llPackID, lcolPacks

'    llPackID = 0
'    Engine.Package.ClearPackage()
'    Engine.Package.Pack_Name = vsPackName
'    If CLng(vlProcID) = 0 Then
'        Engine.Package.Proc_ID = WFCurrentCase.Proc_ID
'    Else
'        Engine.Package.Proc_ID = CLng(vlProcID)
'    End If

'    lcolPacks = Engine.Package.GetPackage
'    If lcolPacks.Count > 0 Then llPackID = lcolPacks.Item(1).Pack_ID
'    lcolPacks = Nothing

'    flGetPackID = llPackID
'End Function

'**********************************************
'*** SUB VOEG COMMENTAAR TOE ***
'**********************************************
'*** BIJ VERLATEN


''**********************************************
''*** SUB VoegCommentaarToe ***
''**********************************************
'Sub VoegCommentaarToe(ByVal vsPersoon, ByVal vsCommentaar, ByVal vsStap)
'    Dim lsCommentaar, loTable, llRowID, lsStap

'    ' Commentaar ophalen
'    lsCommentaar = vsCommentaar
'    If Len(lsCommentaar) > 0 Then
'        ' Tabel ophalen
'        loTable = WFGetProperty("Overzicht Commentaren")
'        ' Nieuwe rij aanmaken
'        llRowID = loTable.CreateRow()
'        ' Opvullen Data
'        Call WFSetProperty("Datum", Day(Now()) & "/" & Month(Now()) & "/" & Year(Now()) & " " & Time, True, llRowID, loTable.Prop_id)
'        Call WFSetProperty("Persoon", vsPersoon, True, llRowID, loTable.Prop_id)
'        Call WFSetProperty("Commentaar", lsCommentaar, True, llRowID, loTable.Prop_id)
'        Call WFSetProperty("Stap", vsStap, True, llRowID, loTable.Prop_id)
'    End If
'End Sub

'******************************
'GLOBALE FUNCTIES
'******************************

'**********************************
'*** SUB SaceCase ***
'**********************************
'Sub Save()
'    ''17-08-2009
'    WFCurrentCase.updatelistinfo()
'    WFCurrentCase.SaveCase() '= Save all values of the properties of the currentcase.
'    Engine.SetComplete() '= Commit of all changes. SetAbort to rollback.
'End Sub


'Sub WFSetDossierNaamAdviesvragen(ByVal WFCurrentCase As Doma.Library.Routing.cCase)
'    Dim lsDossierNaam, lsjaartal As String
'    Dim lsTypeBestuur, lsBestuurNaam, lsTypeOrganisatie As String

'    lsTypeBestuur = WFCurrentCase.GetProperty("type bestuur").ToString
'    lsTypeOrganisatie = WFCurrentCase.GetProperty("type vraagsteller").ToString
'    lsBestuurNaam = WFCurrentCase.GetProperty("Naam bestuur vraagsteller").ToString
'    'lsjaartal = Year(WFCurrentCase.Case_DateStart)
'    lsjaartal = Year(CDate(WFCurrentCase.CaseData.Creation_Date)).ToString

'    lsDossierNaam = lsjaartal & " - " & WFCurrentCase.GetProperty("S_Dossiernummer").ToString & " Adviesvraag"
'    lsDossierNaam = lsDossierNaam & " " & lsTypeOrganisatie
'    lsDossierNaam = lsDossierNaam & ": " & lsTypeBestuur
'    lsDossierNaam = lsDossierNaam & "  " & lsBestuurNaam
'    WFCurrentCase.Case_Name = lsDossierNaam
'End Sub


'Function PadDigits(ByVal lsValue As String, ByVal litotalDigits As Integer) As String
'    'PadDigits =  Right(String(litotalDigits,"0") & lsValue, litotalDigits) 
'    Return lsValue.PadLeft(litotalDigits)
'End Function


'**********************************************
'*** SUB VOEG COMMENTAAR TOE ***
'**********************************************
'*** BIJ VERLATEN

'Sub voegtoe(ByRef WFCurrentCase As Doma.Library.Routing.cCase)
'    If WFCurrentCase.GetProperty("commentaarveld").ToString <> "" Then
'        ' Commentaar van huidige uitvoerder toevoegen aan Overzicht Commentaren
'        'Call VoegCommentaarToe(WFCurrentCase.User_StepExecutor, WFGetProperty("commentaarveld"), WFGetStepName)
'        Call VoegCommentaarToe(WFCurrentCase.StepExecutor, WFCurrentCase.GetProperty("commentaarveld").ToString, WFGetStepName)
'        ' Commentaar leegmaken
'    End If
'    Call WFCurrentCase.SetProperty("commentaarveld", "")
'End Sub


''**********************************************
''*** SUB VoegCommentaarToe ***
''**********************************************
'Sub VoegCommentaarToe(ByRef WFCurrentCase As Doma.Library.Routing.cCase, ByVal vsPersoon As String, ByVal vsCommentaar As String, ByVal vsStap As String)

'    Dim lsCommentaar, loTable, llRowID, lsStap As String
'    ' Commentaar ophalen
'    lsCommentaar = vsCommentaar
'    If Len(lsCommentaar) > 0 Then
'        ' Tabel ophalen
'        loTable = WFCurrentCase.GetProperty("Overzicht Commentaren").ToString
'        ' Nieuwe rij aanmaken
'        llRowID = loTable.CreateRow()
'        ' Opvullen Data
'        'Call WFSetProperty("Datum",Day(Now()) & "/" & Month(Now()) & "/" & Year(Now()) & " " & Time ,True, llRowID, loTable.Prop_id)
'        Call WFCurrentCase.SetProperty("Persoon", vsPersoon, True, llRowID, loTable.Prop_id)
'        Call WFCurrentCase.SetProperty("Commentaar", lsCommentaar, True, llRowID, loTable.Prop_id)
'        Call WFCurrentCase.SetProperty("Stap", vsStap, True, llRowID, loTable.Prop_id)
'    End If
'End Sub

'Sub HaalTrefwoordOp(ByVal lsft_cid As String, ByRef lsTrefwoord As String)
'    Dim lsSQL, goCN, goRS, lsError As String
'    lsError = " "

'    Call sAddToLog("Executing statement: " & lsSQL, WFCurrentCase.StepExecutor)

'    lsSQL = "select TREFWOORD from BB_TREFWOORDEN  where FT_CID=" & lsft_cid
'    Dim lsTempStr As String = ""
'    Using loQuery As Arco.Server.DataQuery = New Arco.Server.DataQuery
'        loQuery.Query = lsSQL
'        Try
'            loQuery.Connect()
'            Using loReader As Server.SafeDataReader = loQuery.ExecuteReader
'                If loReader.Read Then
'                    lsTrefwoord = CheckEmpty(CStr(loReader("TREFWOORD")))
'                End If
'            End Using
'        Catch ex As Exception
'            Arco.Utils.Logging.Log("error HaalTrefwoordOp:" & ex.Message)
'        End Try
'    End Using
'    Call sAddToLog("Executed  statement: " & lsSQL, WFCurrentCase.StepExecutor)

'End Sub

'Sub VoegTrefwoordToe(ByRef WFCurrentCase As Doma.Library.Routing.cCase, ByVal vsTrefwoord As String)
'    Dim lsTrefwoord, loTable, llRowID
'    lsTrefwoord = vsTrefwoord
'    If Len(lsTrefwoord) > 0 Then
'        ' Tabel ophalen
'        loTable = WFGetProperty("trefwoordenlijst")
'        ' Nieuwe rij aanmaken
'        llRowID = loTable.CreateRow()
'        ' Debug.PrintToFile "llRowID :" & llRowID 
'        ' Opvullen Data
'        Call WFSetProperty("trefwoord", lsTrefwoord, True, llRowID, loTable.Prop_id)
'    End If
'End Sub


'Function fsGetType(ByVal lstypebestuur As String) As String
'    Dim lssql, lsBestuur As String
'    lsBestuur = ""
'    lssql = "SELECT type FROM bb_type where meldingstype='" & lstypebestuur & "'"

'    Call sAddToLog("Executing statement: " & lssql, WFCurrentCase.StepExecutor)
'    Using loQuery As Arco.Server.DataQuery = New Arco.Server.DataQuery
'        loQuery.Query = lssql
'        Try
'            loQuery.Connect()
'            Using loReader As Server.SafeDataReader = loQuery.ExecuteReader
'                If loReader.Read Then
'                    If Not IsNothing(loReader("type")) Then
'                        lsBestuur = CStr(loReader("type"))
'                    End If
'                End If
'            End Using
'        Catch ex As Exception
'            Arco.Utils.Logging.Log("error fsGetType:" & ex.Message)
'        End Try
'    End Using

'    Arco.Utils.Logging.Log("Bestuur:" & lsBestuur)
'    Return lsBestuur
'    Call sAddToLog("Executed  statement: " & lssql, WFCurrentCase.StepExecutor)

'End Function


'Function GetNaamBestuur(ByVal libestuurid As Integer) As String
'    Dim lssql, lsBestuurnaam As String
'    lsBestuurnaam = ""

'    Call sAddToLog("Executing statement: " & lssql, WFCurrentCase.StepExecutor)
'    lssql = "SELECT naam FROM bb_adresbesturen where ft_cid='" & libestuurid & "'"
'    Using loQuery As Arco.Server.DataQuery = New Arco.Server.DataQuery
'        loQuery.Query = lssql
'        Try
'            loQuery.Connect()
'            Using loReader As Server.SafeDataReader = loQuery.ExecuteReader
'                If loReader.Read Then
'                    If Not IsNothing(loReader("naam")) Then
'                        lsBestuurnaam = CStr(loReader("naam"))
'                    End If
'                End If
'            End Using
'        Catch ex As Exception
'            Arco.Utils.Logging.Log("error GetNaamBestuur:" & ex.Message)
'        End Try
'    End Using
'    Return lsBestuurnaam
'    Call sAddToLog("Executed  statement: " & lssql, WFCurrentCase.StepExecutor)

'End Function

'***********************************************************************************************************************

'Function GetDienstTeam(ByVal lsbehandelaar As String, ByRef lsteam As String, ByRef lsrole_id As String) As Boolean ' inkomend veld 
'    Dim lssql As String  'uitgaand veld
'    Dim lbOK As Boolean = False

'    lsteam = ""
'    lssql = "Select a.role_name as rolnaam, a.role_id  from rtrole a, rtrolemember b where a.role_id = b.role_id And b.member = '" & lsbehandelaar & "' order by role_description"

'    Call sAddToLog("Executing statement: " & lssql, WFCurrentCase.StepExecutor)
'    Using loQuery As Arco.Server.DataQuery = New Arco.Server.DataQuery
'        loQuery.Query = lssql
'        Try
'            loQuery.Connect()
'            Using loReader As Server.SafeDataReader = loQuery.ExecuteReader
'                If loReader.Read Then
'                    If Not IsNothing(loReader("rolnaam")) Then
'                        lsteam = CStr(loReader("rolnaam"))
'                    End If
'                    If Not IsNothing(loReader("rolID")) Then
'                        lsrole_id = CStr(loReader("rolID"))
'                    End If
'                    lbOK = True
'                End If
'            End Using
'        Catch ex As Exception
'            Arco.Utils.Logging.Log("error GetDienstTeam:" & ex.Message)
'        End Try
'    End Using

'    Call sAddToLog("Executed  statement: " & lssql, WFCurrentCase.StepExecutor)
'    Return lbOK

'End Function


'************************************************************************************************************************
'Function GetAfdeling(ByVal lsrole_id As String) As String
'    Dim lssql, lsAfdeling As String
'    lsAfdeling = ""

'    lssql = " SELECT  ROLE_NAME ,ROLE_DESCRIPTION,ROLE_STRUCTURED FROM RTROLE WHERE ROLE_STRUCTURED=1  AND  "
'    lssql = lssql & "ROLE_ID IN (SELECT PARENT_ROLE FROM RTROLE_LINKS WHERE RTROLE_LINKS.CHILD_ROLE='" & lsrole_id & "') union "
'    lssql = lssql & "SELECT  ROLE_NAME ,ROLE_DESCRIPTION,ROLE_STRUCTURED FROM RTROLE WHERE ROLE_STRUCTURED=1  AND "
'    lssql = lssql & " ROLE_ID IN (SELECT PARENT_ROLE FROM RTROLE_LINKS WHERE RTROLE_LINKS.CHILD_ROLE IN (SELECT PARENT_ROLE FROM RTROLE_LINKS WHERE RTROLE_LINKS.CHILD_ROLE='"&lsrole_id&"')) union "
'    lssql = lssql & " SELECT   ROLE_NAME ,ROLE_DESCRIPTION,ROLE_STRUCTURED FROM RTROLE WHERE ROLE_STRUCTURED=1  AND "
'    lssql = lssql & "ROLE_ID IN (SELECT PARENT_ROLE FROM RTROLE_LINKS WHERE RTROLE_LINKS.CHILD_ROLE IN "
'    lssql = lssql & " (SELECT PARENT_ROLE FROM RTROLE_LINKS WHERE RTROLE_LINKS.CHILD_ROLE IN (SELECT PARENT_ROLE FROM RTROLE_LINKS WHERE RTROLE_LINKS.CHILD_ROLE='"&lsrole_id&"'))) "
'    lssql = lssql & "ORDER BY role_description "

'    Call sAddToLog("Executing statement: " & lssql, WFCurrentCase.StepExecutor)
'    Using loQuery As Arco.Server.DataQuery = New Arco.Server.DataQuery
'        loQuery.Query = lssql
'        Try
'            loQuery.Connect()
'            Using loReader As Server.SafeDataReader = loQuery.ExecuteReader
'                If loReader.Read Then
'                    If Not IsNothing(loReader("ROLE_NAME")) Then
'                        lsAfdeling = CStr(loReader("ROLE_NAME"))
'                    End If
'                End If
'            End Using
'        Catch ex As Exception
'            Arco.Utils.Logging.Log("error GetAfdeling:" & ex.Message)
'        End Try
'    End Using

'    Return lsAfdeling
'    Call sAddToLog("Executed  statement: " & lssql, WFCurrentCase.StepExecutor)

'End Function

'Function UpdateTableRecords(ByVal liCase_id As Integer) As Boolean
'    Dim lsSQL1 As String
'    Dim lbok1 As Boolean = False
'    lsSQL1 = "update rtcase_property  set dm_object_id=0 where prop_id in  ("
'    lsSQL1 = lsSQL1 & " select prop_id from rtproperty where parent_prop_id  in ("
'    lsSQL1 = lsSQL1 & " select prop_id from rtproperty where prop_type='TABLE'"
'    lsSQL1 = lsSQL1 & " and proc_id = (select proc_id from rtprocedure where proc_name='Adviesvragen')))"
'    lsSQL1 = lsSQL1 & " and CASE_ID='" & liCase_id & "'"

'    Call sAddToLog("Executing statement: " & lsSQL1, WFCurrentCase.StepExecutor)
'    Using loQuery As Arco.Server.DataQuery = New Arco.Server.DataQuery
'        loQuery.Query = lsSQL1
'        Try
'            loQuery.Connect()
'            loQuery.ExecuteNonQuery(lsSQL1)
'            lbok1 = True
'        Catch ex As Exception
'            Arco.Utils.Logging.Log("error UpdateTableRecords:" & ex.Message)
'            lbok1 = False
'        End Try
'    End Using
'    Call sAddToLog("Executed  statement: " & lsSQL1, WFCurrentCase.StepExecutor)
'    Return lbok1

'End Function

''**********************************************
''*** HaalContactpersoon
''**********************************************

'Sub HaalContactpersoon(ByVal lsft_cid)

'    'select naam, voornaam , straatnr, postcode , gemeente from BB_klagers where ft_cid=2020'
'    Dim lsSQL, goCN, goRS, lsError
'    lsError = " "
'    lsSQL = "select naam, voornaam , straatnr, postcode , gemeente,email from BB_klagers where FT_CID=" & lsft_cid
'    Call AddToLog("Executing statement: " & lsSQL)
'    goCN = CreateObject("adodb.connection")
'    goCN.open(gsDSN, gsUser, gsPassword)

'    If goCN.state = 1 Then
'        goRS = CreateObject("adodb.recordset")
'        goRS.open(lsSQL, goCN)
'        If Not goRS.eof Then

'            Call WFSetProperty("klager_naam", goRS("naam"))
'            Call WFSetProperty("klager_voornaam", goRS("voornaam"))
'            Call WFSetProperty("klager_straatnr", goRS("straatnr"))
'            Call WFSetProperty("klager_postnummer", goRS("postcode"))
'            Call WFSetProperty("klager_gemeente", goRS("gemeente"))

'            If Not IsNull(goRS("email")) Then


'                Call WFSetProperty("klager_email", goRS("email"))
'            End If
'            ' terug leegmaken anders terug opgehaald
'            Call WFSetProperty("klager ID", "")
'        End If
'        goRS.close()
'        goRS = Nothing
'    Else
'        'lserror="fout DB, verkeerde status"
'    End If
'    goCN.close()
'    goCN = Nothing
'    Call AddToLog("Executing statement: " & lsSQL)
'End Sub



''**********************************************
''*** Function HaalBestuurOp
''**********************************************
'Sub HaalBestuurOp(ByVal lsft_cid)
'    Dim lsSQL, goCN, goRS, lsError
'    lsError = " "
'    lsSQL = "select NAAM,STRAATNR, POSTCODE, GEMEENTE from BB_ADRESBESTUREN  where FT_CID=" & lsft_cid
'    Call AddToLog("Executing statement: " & lsSQL)
'    goCN = CreateObject("adodb.connection")
'    goCN.open(gsDSN, gsUser, gsPassword)
'    If goCN.state = 1 Then
'        goRS = CreateObject("adodb.recordset")
'        goRS.open(lsSQL, goCN)
'        If Not goRS.eof Then

'            Call WFSetProperty("bestuur_naam", goRS("naam"))
'            Call WFSetProperty("bestuur_straatnr", goRS("straatnr"))
'            Call WFSetProperty("bestuur_postnummer", goRS("postcode"))
'            Call WFSetProperty("bestuur_gemeente", goRS("gemeente"))

'            ' terug leegmaken anders terug opgehaald
'            Call WFSetProperty("bestuur", "")
'        End If
'        goRS.close()
'        goRS = Nothing
'    Else
'        'lserror="fout DB, verkeerde status"
'    End If
'    goCN.close()
'    goCN = Nothing
'    Call AddToLog("Executing statement: " & lsSQL)
'End Sub


''**********************************************
''*** Function HaalTermijnVanType
'' *** => haalt het aantal dagen op van een bepaalde type in  BB_TYPE,return type integer
''**********************************************
'Function HaalTermijnVanType(ByVal lsType)
'    'select termijn from BB_type where meldingstype = 'Gemeente'
'    Dim lsSQL, goCN, goRS
'    Dim lstemp
'    lsSQL = "select termijn from BB_type where meldingstype = '" & lsType & "'"
'    Call AddToLog("Executing statement: " & lsSQL)
'    goCN = CreateObject("adodb.connection")
'    goCN.open(gsDSN, gsUser, gsPassword)

'    If goCN.state = 1 Then
'        goRS = CreateObject("adodb.recordset")
'        goRS.open(lsSQL, goCN)
'        If Not goRS.eof Then

'            lstemp = goRS("termijn")
'            If lstemp <> "" Then
'                lserror = CInt(lstemp)
'            End If

'        End If
'        goRS.close()
'        goRS = Nothing
'    Else
'        lstemp = 0
'    End If
'    goCN.close()
'    goCN = Nothing
'    Call AddToLog("Executed  statement: " & lsSQL)
'    HaalTermijnVanType = lstemp
'End Function


'Sub HaalBestuurOp2(ByVal lsft_cid As String, ByRef lsHBType As String, ByRef lsHBNaam As String, ByRef lsHBNIS As String, ByRef lsHBStraatnr As String, _
'                      ByRef lsHBPostcode As String, ByRef lsHBGemeente As String, ByVal WFCurrentCase As Doma.Library.Routing.cCase)
'    Dim lsSQL, goCN, goRS, lsError As String

'    Call sAddToLog("Executing statement: " & lsSQL, WFCurrentCase.StepExecutor)
'    lsError = " "
'    lsSQL = "select TYPE, NAAM,NIS,STRAATNR, POSTCODE, GEMEENTE from BB_ADRESBESTUREN  where FT_CID=" & lsft_cid
'    Using loQuery As Arco.Server.DataQuery = New Arco.Server.DataQuery
'        loQuery.Query = lsSQL
'        Try
'            loQuery.Connect()
'            Using loReader As Server.SafeDataReader = loQuery.ExecuteReader
'                While loReader.Read
'                    lsHBType = CheckEmpty(CStr(loReader("type")))
'                    lsHBNaam = CheckEmpty(CStr(loReader("naam")))
'                    lsHBNIS = CheckEmpty(CStr(loReader("NIS")))
'                    lsHBStraatnr = CheckEmpty(CStr(loReader("straatnr")))
'                    lsHBPostcode = CheckEmpty(CStr(loReader("postcode")))
'                    lsHBGemeente = CheckEmpty(CStr(loReader("gemeente")))
'                End While
'            End Using
'        Catch ex As Exception
'            Arco.Utils.Logging.Log("error HaalBestuurOp2:" & ex.Message)
'        End Try
'    End Using
'    Call sAddToLog("Executed  statement: " & lsSQL, WFCurrentCase.StepExecutor)

'End Sub


'Function CheckEmpty(ByVal lsValue As String) As String
'    'If IsNull(lsValue) Then
'    '    lsValue = ""
'    'End If
'    If String.IsNullOrEmpty(lsValue) Then
'        lsValue = ""
'    End If
'    CheckEmpty = lsValue
'End Function

'Sub HaalContactpersoon2(ByVal lsft_cid As String, ByRef lsCPNaam As String, ByRef lsCPvoornaam As String, ByRef lsCPStraatnr As String, _
'                           ByRef lsCPPostnr As String, ByRef lsCPGemeente As String, ByRef lsCPEmail As String, ByRef lsCPTelefoon As String, _
'                           ByRef lsCPfax As String, ByVal WFCurrentCase As Doma.Library.Routing.cCase)
'    'select naam, voornaam , straatnr, postcode , gemeente from BB_klagers where ft_cid=2020'
'    Dim lsSQL, lsError As String
'    Dim goCN, gORS As Object
'    lsError = " "

'    lsError = " "
'    lsSQL = "select naam, voornaam , straatnr, postcode , gemeente,email, telefoon,fax from BB_klagers where FT_CID=" & lsft_cid
'    Using loQuery As Arco.Server.DataQuery = New Arco.Server.DataQuery
'        loQuery.Query = lsSQL
'        Try
'            loQuery.Connect()
'            Using loReader As Server.SafeDataReader = loQuery.ExecuteReader
'                If loReader.Read Then
'                    lsCPNaam = CheckEmpty(CStr(loReader("naam")))
'                    lsCPvoornaam = CheckEmpty(CStr(loReader("voornaam")))
'                    lsCPStraatnr = CheckEmpty(CStr(loReader("straatnr")))
'                    lsCPPostnr = CheckEmpty(CStr(loReader("postcode")))
'                    lsCPGemeente = CheckEmpty(CStr(loReader("gemeente")))
'                    lsCPEmail = CheckEmpty(CStr(loReader("email")))
'                    lsCPTelefoon = CheckEmpty(CStr(loReader("telefoon")))
'                    lsCPfax = CheckEmpty(CStr(loReader("fax")))
'                End If
'            End Using
'        Catch ex As Exception
'            Arco.Utils.Logging.Log("error HaalContactpersoon2:" & ex.Message)
'        End Try
'    End Using

'    Call sAddToLog("Executed  statement: " & lsSQL, wfCurrentCase.StepExecutor)

'End Sub

'Function GetUsersFromRole(ByVal lsRoleName As String, ByVal sCurrentUser As String) As String
'    Dim lsUsers, lsUserName As String
'    Dim lsSQL, lsTemp As String
'    Dim lbOK As Boolean

'    Call sAddToLog("Executing statement: " & lsSQL, sCurrentUser)

'    lsSQL = "SELECT distinct Member, user_display_name FROM RTRoleMember,rtusers where rtusers.user_login = replace(member,'_DATABASE\','') "
'    lsSQL = lsSQL & "and rtusers.user_Desc like '%V%' and Role_ID =" & flGetRoleId(Replace(lsRoleName, "(Role) ", "")) & " ORDER BY user_display_name"
'    Dim lsTempStr As String = ""
'    Using loQuery As Arco.Server.DataQuery = New Arco.Server.DataQuery
'        loQuery.Query = lsSQL
'        Try
'            loQuery.Connect()
'            Using loReader As Server.SafeDataReader = loQuery.ExecuteReader
'                While loReader.Read
'                    lsTempStr = CheckEmpty(CStr(loReader("Member")))
'                    lsTemp = lsTemp & lsTempStr & ", "
'                End While
'            End Using
'        Catch ex As Exception
'            Arco.Utils.Logging.Log("error GetUsersFromRole:" & ex.Message)
'        End Try
'    End Using
'    Call sAddToLog("Executed  statement: " & lsSQL, sCurrentUser)

'    'lsTemp = fsReplaceTo(lsTemp)
'    If lsTemp <> "" Then
'        GetUsersFromRole = Left(lsTemp, Len(lsTemp) - 2)
'    Else
'        GetUsersFromRole = ""
'    End If

'End Function

'Function flGetRoleId(ByVal lsRoleName As String) As Long
'    Dim loCol
'    Dim lsrole_id() As String
'    lsrole_id = 0
'    If lsRoleName <> "" Then
'        Engine.UserManager.Role.Clear()
'        Engine.UserManager.Role.Role_Name = lsRoleName
'        loCol = Engine.UserManager.Role.GetRoles()
'        If loCol.Count > 0 Then
'            lsrole_id = loCol.Item(1).Role_ID
'        End If
'    End If
'    flGetRoleId = lsrole_id
'End Function

'Sub HaalGegevensOpAfdeling(ByRef naam, ByRef straatnr, ByRef postcode, ByRef gemeente, ByRef emailadres, ByRef telefoonnr, ByRef fax)
'    Dim lsSQL, goCN, goRS, lsError
'    lsSQL = "select naam,straatnr,postcode,gemeente,emailadres, telefoonnr,fax from BB_AFDELING where upper(naam) = upper('" & naam & "');"
'    Call AddToLog("Executing statement: " & lsSQL)
'    goCN = CreateObject("adodb.connection")
'    goCN.open(gsDSN, gsUser, gsPassword)
'    If goCN.state = 1 Then
'        goRS = CreateObject("adodb.recordset")
'        goRS.open(lsSQL, goCN)
'        If Not goRS.eof Then
'            straatnr = goRS("straatnr")
'            postcode = goRS("postcode")
'            gemeente = goRS("gemeente")
'            emailadres = goRS("emailadres")
'            emailadres = goRS("telefoonnr")
'            fax = goRS("fax")
'        End If
'        goRS.close()
'        goRS = Nothing
'    End If
'    goCN.close()
'    goCN = Nothing
'    Call AddToLog("Executed  statement: " & lsSQL)
'End Sub

'#Region "DBFuncties"
''() DB FUNCTIES



'Sub HaalContactpersoon2(ByVal lsft_cid)

'    'select naam, voornaam , straatnr, postcode , gemeente from BB_klagers where ft_cid=2020'
'    Dim lsSQL, goCN, goRS, lsError
'    lsError = " "
'    lsSQL = "select naam, voornaam , straatnr, postcode , gemeente,email from BB_klagers where FT_CID=" & lsft_cid
'    Call AddToLog("Executing statement: " & lsSQL)
'    goCN = CreateObject("adodb.connection")
'    goCN.open(gsDSN, gsUser, gsPassword)

'    If goCN.state = 1 Then
'        goRS = CreateObject("adodb.recordset")
'        goRS.open(lsSQL, goCN)
'        If Not goRS.eof Then

'            Call WFSetProperty("klager_naam", goRS("naam"))
'            Call WFSetProperty("klager_voornaam", goRS("voornaam"))
'            Call WFSetProperty("klager_straatnr", goRS("straatnr"))
'            Call WFSetProperty("klager_postnummer", goRS("postcode"))
'            Call WFSetProperty("klager_gemeente", goRS("gemeente"))

'            If Not IsNull(goRS("email")) Then
'                Call WFSetProperty("klager_email", goRS("email"))
'            End If
'            ' terug leegmaken anders terug opgehaald
'            Call WFSetProperty("klager ID", "")
'        End If
'        goRS.close()
'        goRS = Nothing
'    Else
'        'lserror="fout DB, verkeerde status"
'    End If
'    goCN.close()
'    goCN = Nothing
'    Call AddToLog("Executed  statement: " & lsSQL)
'End Sub


'Sub Haalanderelijst(ByVal lsft_cid, ByRef vsbeslissingsorgaan, ByRef vsdatum_besluit, ByRef vstitel_besluit)
'    Dim lsSQL, goCN, goRS, lsError
'    lsError = " "
'    lsSQL = "select ft_Cid , beslissingsorgaan , datum_besluit,titel_besluit from BB_ander_besl where ft_Cid=" & lsft_cid
'    Call AddToLog("Executing statement: " & lsSQL)
'    goCN = CreateObject("adodb.connection")
'    goCN.open(gsDSN, gsUser, gsPassword)

'    If goCN.state = 1 Then
'        goRS = CreateObject("adodb.recordset")
'        goRS.open(lsSQL, goCN)
'        If Not goRS.eof Then
'            vsbeslissingsorgaan = goRS("beslissingsorgaan")
'            vsdatum_besluit = goRS("datum_besluit")
'            vstitel_besluit = goRS("titel_besluit")
'        End If
'        goRS.close()
'        goRS = Nothing
'    Else
'        'lserror="fout DB, verkeerde status"
'    End If
'    goCN.close()
'    goCN = Nothing
'    Call AddToLog("Executed  statement: " & lsSQL)
'End Sub


'Sub Haalinzendingsplichtigelijst(ByVal lsft_cid, ByRef vsdatum_besluit, ByRef vssoort_besluit, ByRef vspost_datum, ByRef vsdatum_in, ByRef vsInitTermijn)

'    'on error resume Next

'    Dim lsSQL, goCN, goRS, lsError
'    lsError = " "
'    lsSQL = "select FT_CID , datum_besluit, soort_besluit,post_datum ,datum_in, initiele_termijn  from BB_inzend_besl where ft_Cid=" & lsft_cid
'    Call AddToLog("Executing statement: " & lsSQL)

'    goCN = CreateObject("adodb.connection")
'    goCN.open(gsDSN, gsUser, gsPassword)
'    If goCN.state = 1 Then
'        goRS = CreateObject("adodb.recordset")
'        goRS.open(lsSQL, goCN)
'        If Not goRS.eof Then
'            vsdatum_besluit = goRS("datum_besluit")
'            vssoort_besluit = goRS("soort_besluit")
'            vspost_datum = goRS("post_datum")
'            vsdatum_in = goRS("datum_in")
'            vsInitTermijn = goRS("INITIELE_TERMIJN")
'        End If
'        goRS.close()
'        goRS = Nothing
'    Else
'        'lserror="fout DB, verkeerde status"
'    End If
'    goCN.close()
'    goCN = Nothing
'    Call AddToLog("Executed  statement: " & lsSQL)
'End Sub


'Sub Haalmeldingslijst(ByVal lsft_cid, ByRef vsgemeente, ByRef vsdatumzitting, ByRef vspostdatum, ByRef vsontvangstdatum, ByRef vsInitTermijn)
'    'select termijn from BB_type where meldingstype = 'Gemeente'

'    Dim lsSQL, goCN, goRS, lsError
'    lsError = " "
'    lsSQL = "select GEMEENTE, DATUM_ZITTING,POST_DATUM , ONTVANGSTDATUM,INITIELE_TERMIJN  from bb_meldingslijst where FT_CID=" & lsft_cid
'    Call AddToLog("Executing statement: " & lsSQL)
'    goCN = CreateObject("adodb.connection")
'    goCN.open(gsDSN, gsUser, gsPassword)
'    'goCN.open "ARCO" ,"arco_appl" , "arco_appl"

'    If goCN.state = 1 Then
'        goRS = CreateObject("adodb.recordset")
'        goRS.open(lsSQL, goCN)
'        If Not goRS.eof Then
'            vsgemeente = goRS("gemeente")
'            vsdatumzitting = goRS("datum_zitting")
'            vspostdatum = goRS("post_datum")
'            vsontvangstdatum = goRS("ontvangstdatum")
'            vsInitTermijn = goRS("INITIELE_TERMIJN")
'            If vsgemeente = "" And vsdatumzitting = "" And vspostdatum = "" Then
'                vsgemeente = ""
'                vsdatumzitting = ""
'                vspostdatum = ""
'                vsontvangstdatum = ""
'                vsInitTermijn = ""
'            End If
'        End If
'        goRS.close()
'        goRS = Nothing
'    Else
'        'lserror="fout DB, verkeerde status"
'    End If
'    goCN.close()
'    goCN = Nothing
'    Call AddToLog("Executed statement: " & lsSQL)
'End Sub


'Function HaalAfdelingOpVanGebruiker(ByVal lsUser)
'    Dim lsafdeling
'    lsafdeling = ""

'    Dim lsSQL, goCN, goRS, lsError
'		lsSQL = "select rtrole.role_name from rtrolemember,rtrole where rtrolemember.member=UPPER('"&lsuser&"')and role_name like '%Afdeling%' and rtrole.role_id= rtrolemember.role_ID and rownum <2"
'    Call AddToLog("Executing statement: " & lsSQL)
'    goCN = CreateObject("adodb.connection")
'    goCN.open(gsDSN, gsUser, gsPassword)
'    'goCN.open "ARCO" ,"arco_appl" , "arco_appl"
'    If goCN.state = 1 Then
'        goRS = CreateObject("adodb.recordset")
'        goRS.open(lsSQL, goCN)
'        If Not goRS.eof Then
'            lsafdeling = goRS("role_name")
'        End If
'        goRS.close()
'        goRS = Nothing
'    End If
'    goCN.close()
'    goCN = Nothing
'    Call AddToLog("Executed  statement: " & lsSQL)
'    HaalAfdelingOpVanGebruiker = lsafdeling
'End Function




'Sub SelectTypeBestuur()
'    Dim lslijstbesluitnr
'    lslijstbesluitnr = WFGetProperty("lijstbesluit_nr")
'    If lslijstbesluitnr <> "" Then
'        'select BB_meldingslijst.type from  BB_meldingslijst where BB_meldingslijst.ft_cid = 2108
'        Dim lsSQL, goCN, goRS, lsError
'        lsSQL = "select BB_meldingslijst.type from  BB_meldingslijst where BB_meldingslijst.ft_cid =" & lslijstbesluitnr
'        Call AddToLog("Executing statement: " & lsSQL)
'        goCN = CreateObject("adodb.connection")
'        goCN.open(gsDSN, gsUser, gsPassword)
'        'goCN.open "ARCO" ,"arco_appl" , "arco_appl"
'        If goCN.state = 1 Then
'            goRS = CreateObject("adodb.recordset")
'            goRS.open(lsSQL, goCN)
'            If Not goRS.eof Then
'                Call WFSetProperty("type bestuur", goRS("type"))
'            End If
'            goRS.close()
'            goRS = Nothing
'        End If
'        goCN.close()
'        goCN = Nothing
'        Call AddToLog("Executed  statement: " & lsSQL)
'    End If
'End Sub
'#End Region

'Private Sub s00_ToonWelkeMailbox()
'    Call AddToLog("BEGIN (00) Toon welke mailbox")
'    '(00) Toon welke mailbox

'    If WFGetProperty("Medium adviesvraag") = "mail" Then
'        Call WFSetPropertyVisible("via_welke_mailbox", True)
'    Else
'        Call WFSetPropertyVisible("via_welke_mailbox", False)
'        If WFGetProperty("Medium adviesvraag") <> "" Then
'            Call WFSetProperty("via_welke_mailbox", "")
'        Else
'        End If
'    End If
'    Engine.SetComplete()
'    Call AddToLog("EINDE (00) Toon welke mailbox")
'End Sub

''**********************************
''*** FUNCTIE  fsUserData ***
''**********************************
'Function fsUserData(ByVal vsUsername As String, ByVal vsWhat As String) As String
'    Dim lcolUsers
'    fsUserData = ""
'    vsUsername = Trim(vsUsername)
'    If InStr(vsUsername, "\") > 0 Then
'        Engine.UserManager.USERS.Clear()
'        'Strip het voorvoegsel "_DATABASE\", aangezien dit niet in het login veld staat van de database
'        If Left(UCase(vsUsername), "9") = "_DATABASE" Then vsUsername = Mid(vsUsername, InStr(vsUsername, "\") + 1)
'        Engine.UserManager.USERS.USER_LOGIN = vsUsername
'        lcolUsers = Engine.UserManager.USERS.GetUSERS
'        If lcolUsers.Count = 1 Then
'            Select Case UCase(vsWhat)
'                Case "USER_DISPLAY_NAME"
'                    fsUserData = lcolUsers.Item(1).USER_DISPLAY_NAME
'                Case "USER_MAIL"
'                    fsUserData = lcolUsers.Item(1).USER_MAIL
'                Case "USER_PHONE"
'                    fsUserData = lcolUsers.Item(1).USER_PHONE
'                Case "USER_DISTR_NOTIF"
'                    fsUserData = lcolUsers.Item(1).USER_DISTR_NOTIF
'            End Select
'        End If
'        lcolUsers = Nothing
'    Else
'        Engine.UserManager.USERS.Clear()
'        Engine.UserManager.USERS.USER_LOGIN = vsUsername
'        lcolUsers = Engine.UserManager.USERS.GetUSERS
'        If lcolUsers.Count = 1 Then
'            Select Case UCase(vsWhat)
'                Case "USER_DISPLAY_NAME"
'                    fsUserData = lcolUsers.Item(1).USER_DISPLAY_NAME
'                Case "USER_MAIL"
'                    fsUserData = lcolUsers.Item(1).USER_MAIL
'                Case "USER_PHONE"
'                    fsUserData = lcolUsers.Item(1).USER_PHONE
'                Case "USER_DISTR_NOTIF"
'                    fsUserData = lcolUsers.Item(1).USER_DISTR_NOTIF
'            End Select
'        End If
'        lcolUsers = Nothing
'    End If
'End Function

