<%
Class MTLog

	Private strSql, strExtra, aryConfig, aryDB
	Private strDatabaseType, strTablePrefix
	Private blnEnableLog, strIPExclude
	Private strQuerystringFilter, strDefaultDoc
	Private strPageTitle, intTimeOffset

	Public Property Let Database(pDatabase)
		aryDB 			= pDatabase
		strDatabaseType = aryDB(0)
		strTablePrefix 	= aryDB(5)
	End Property

	Public Property Let Config(pConfig)
		aryConfig 				= pConfig
		blnEnableLog			= aryConfig(intMTEnableLog)
		strIPExclude 			= aryConfig(intMTIPExclude)
		strQuerystringFilter 	= aryConfig(intMTQuerystringFilter)
		strDefaultDoc 			= aryConfig(intMTDefaultDoc)
		intTimeOffset			= aryConfig(intMTTimeOffset)
	End Property

	Public Property Let PageTitle(pPageTitle)
		strPageTitle = pPageTitle
		strPageTitle = Left(Trim(strPageTitle), 100)
	End Property

	Public ActualUrl

	Public Sub LogFile(strLogUrl, intLogType, strResolution)

		intLogType = CInt(intLogType)

		Dim strDateTime : strDateTime 		= DateAdd("h", intTimeOffset, Now())
		Dim strReferrer : strReferrer 		= Request.ServerVariables("HTTP_REFERER")
		Dim strScriptName : strScriptName 	= Request.ServerVariables("SCRIPT_NAME")
		Dim strUserAgent : strUserAgent 	= Request.ServerVariables("HTTP_USER_AGENT")
		Dim strQuerystring : strQuerystring = Request.Querystring
		Dim intSessionID : intSessionID		= Session.SessionID
		Dim strHost : strHost 				= Request.ServerVariables("REMOTE_HOST")
		Dim StrLanguage : strLanguage 		= Request.ServerVariables("HTTP_ACCEPT_LANGUAGE")
		Dim strIPAddress : strIPAddress		= Request.ServerVariables("HTTP_X_FORWARDED_FOR")

		strIPAddress = FormatIPAddress(strIPAddress)
		If strIPAddress = "" Then
			strIPAddress = Request.ServerVariables("REMOTE_ADDR")
		Else
			If IsPrivateIP(strIPAddress) = True Then
				If IsPrivateIP(Request.ServerVariables("REMOTE_ADDR")) = False Then
					strIPAddress = Request.ServerVariables("REMOTE_ADDR")
				End If
			End If
		End If
		
		If IsIPAddress(strHost) = True Then
			strHost = ""
		End If

		Select Case intLogType
		Case 0
			Dim strScriptUrl : strScriptUrl	= strScriptName
			strQuerystring = FilterQuerystring(strQuerystring)
			If strQuerystring <> "" Then
				strScriptUrl	= strScriptUrl & "?" & strQuerystring
			End If
		Case 1
			strScriptName	= ExtractScriptName(strLogUrl)
			strScriptUrl	= strLogUrl
		Case 2
			strScriptName	= ExtractScriptName(strReferrer)
			strQuerystring	= FilterQuerystring(ExtractQuerystring(strReferrer))
			strScriptUrl	= ExtractScriptName(strReferrer)
			If strQuerystring <> "" Then
				strScriptUrl = strScriptUrl & "?" & strQuerystring
			End If
			strReferrer		= strLogUrl
			
			If strDefaultDoc <> "" Then

				If Right(strScriptName, 1) = "/" Then
					strScriptName = strScriptName & strDefaultDoc
				End If

				If Right(strScriptUrl, 1) = "/" Then
					strScriptUrl = strScriptUrl & strDefaultDoc
				ElseIf InStr(strScriptUrl, "/?") > 0 Then
					Dim aryScript : aryScript = Split(strScriptUrl, "/?")
					If UBound(aryScript) = 1 Then
						strScriptUrl = aryScript(0) & "/" & strDefaultDoc & "?" & aryScript(1)
					ElseIf UBound(aryScript) = 0 Then
						strScriptUrl = aryScript(0) & "/" & strDefaultDoc
					End If
				End If
			End If
			
		End Select
		
		Dim blnLogFile
		If strScriptName = "" Then
			blnLogFile = False
		Else
			blnLogFile = True
		End If
		
		If blnLogFile = True Then
		
			Dim strPath : strPath = ExtractPath(strScriptName)
			Dim strExtension : strExtension = ExtractFileType(strScriptName)

			strSql = "SELECT pn_id, pn_url, pn_page, pn_path, pn_extension, pn_label " &_
				"FROM " & strTablePrefix & "PageNames " &_
				"WHERE pn_url = " & FormatDatabaseString(strScriptUrl, 255)
				
			Dim rsUrl : Set rsUrl = Server.CreateObject("ADODB.Recordset")
			
			If strDatabaseType = "MYSQL" Then
				rsUrl.CursorLocation = 3
			End If
			
			rsUrl.Open strSql, objConn, 1, 2, &H0001

			If rsUrl.Eof Then
				rsUrl.AddNew
				rsUrl(1) = ProtectInsert(strScriptUrl, 255)
				rsUrl(2) = ProtectInsert(strScriptName, 255)
				rsUrl(3) = ProtectInsert(strPath, 255)
				rsUrl(4) = ProtectInsert(strExtension, 10)
				rsUrl(5) = ProtectInsert(strPageTitle, 100)
				rsUrl.Update
			ElseIf strPageTitle <> "" Then
				If rsUrl(5) <> strPageTitle Or IsNull(rsUrl(5)) Then
					rsUrl(5) = ProtectInsert(strPageTitle, 100)
					rsUrl.Update
				End If
			End If
			Dim intPage : intPage = rsUrl("pn_id")

			rsUrl.Close : Set rsUrl = Nothing
			
			Dim intIPNumber : intIPNumber = ConvertIPAddressToLong(strIPAddress)
			
			Dim rsDefinitions : Set rsDefinitions = Server.CreateObject("ADODB.Recordset")
			
			strSql = "SELECT d_id, d_name, d_regexp, d_extra, d_type " &_
				"FROM " & strTablePrefix & "Definitions " &_
				"ORDER BY d_id ASC"
			
			rsDefinitions.Open strSql, objConn, 0, 1, &H0001
			
			Dim intUserAgent
			Dim strRobot : strRobot = MatchDefinition(rsDefinitions, strUserAgent, 3)
			
			If strRobot = "" Then
				strSql = "SELECT s_id, s_ip, s_hostname, s_useragent, s_browser, " &_
					"s_os, s_language, s_country, s_screenarea " &_
					"FROM " & strTablePrefix & "Sessions " &_
					"WHERE s_id = " & intSessionID

				Dim rsSession : Set rsSession = Server.CreateObject("ADODB.Recordset")
				rsSession.Open strSql, objConn, 1, 2, &H0001
				
				If rsSession.Eof Then
					
					Dim strCountry
					
					If IsPrivateIP(strIPAddress) = True Then
						strCountry = "00"
					Else
						strCountry = GetCountry(intIPNumber)
					End If
					
					Dim strBrowser : strBrowser = MatchDefinition(rsDefinitions, strUserAgent, 1)
					Dim strOS : strOs = MatchDefinition(rsDefinitions, strUserAgent, 2)
					
					strLanguage = CleanLanguage(strLanguage)
					intUserAgent = CheckName(2, strUserAgent)
					
					Dim intHost : intHost = CheckName(1, strHost)
					Dim intResolution : intResolution = CheckName(3, strResolution)
					Dim intBrowser : intBrowser = CheckName(4, strBrowser)
					Dim intOs : intOs = CheckName(5, strOs)

					rsSession.Addnew
					rsSession(0) = intSessionID
					rsSession(1) = intIPNumber
					rsSession(2) = intHost
					rsSession(3) = intUserAgent
					rsSession(4) = intBrowser
					rsSession(5) = intOs
					rsSession(6) = ProtectInsert(strLanguage, 5)
					rsSession(7) = ProtectInsert(strCountry, 2)
					rsSession(8) = intResolution
					rsSession.Update
				End If
				
				rsSession.Close : Set rsSession = Nothing
				
				Dim intReferrer : intReferrer = 0
				
				If strReferrer <> "" Then
					strSql = "SELECT r_id, r_url, r_rn_id, r_k_id " &_
						"FROM " & strTablePrefix & "Referrers " &_
						"WHERE r_url = " & FormatDatabaseString(strReferrer, 255)

					Dim rsReferrer : Set rsReferrer = Server.CreateObject("ADODB.Recordset")
					
					If strDatabaseType = "MYSQL" Then
						rsReferrer.CursorLocation = 3
					End If
					
					rsReferrer.Open strSql, objConn, 1, 2, &H0001
					
					If rsReferrer.Eof Then
						Dim strReferrerPage : strReferrerPage = ExtractPage(strReferrer)
					
						strSql = "SELECT rn_id, rn_page, rn_host, rn_domain, rn_extension " &_
							"FROM " & strTablePrefix & "ReferrerNames " &_
							"WHERE rn_page = " & FormatDatabaseString(strReferrerPage, 255)

						Dim rsReferrerName : Set rsReferrerName = Server.CreateObject("ADODB.Recordset")
						
						If strDatabaseType = "MYSQL" Then
							rsReferrerName.CursorLocation = 3
						End If
						
						rsReferrerName.Open strSql, objConn, 1, 2, &H0001
						
						If rsReferrerName.Eof Then
							Dim strReferrerHost : strReferrerHost = ExtractHost(strReferrer)
							Dim strReferrerDomain : strReferrerDomain = ExtractDomain(strReferrerHost)
							Dim strReferrerExtension : strReferrerExtension = ExtractExtension(strReferrerDomain)
							rsReferrerName.AddNew
							rsReferrerName(1) = ProtectInsert(strReferrerPage, 255)
							rsReferrerName(2) = ProtectInsert(strReferrerHost, 255)
							rsReferrerName(3) = ProtectInsert(strReferrerDomain, 100)
							rsReferrerName(4) = ProtectInsert(strReferrerExtension, 10)
							rsReferrerName.Update
						End If
						
						Dim intReferrerName : intReferrerName = rsReferrerName(0)
						
						rsReferrerName.Close : Set rsReferrerName = Nothing
						
						Dim intKeywords : intKeywords = 0
						
						If InStr(strReferrer, Request.ServerVariables("SERVER_NAME")) = 0 Then
							
							Dim strSite : strSite = MatchDefinition(rsDefinitions, strReferrer, 4)
							
							If strSite <> "" Then
								Dim strKeywordPrefix : strKeywordPrefix = strExtra
								Dim strKeywords : strKeywords = ExtractKeywords(strReferrer, strKeywordPrefix)
								
								If strKeywords <> "" Then
									
									Dim intSite : intSite = CheckName(8, strSite)
									strSql = "SELECT k_id, k_value, k_site " &_
										"FROM " & strTablePrefix & "Keywords " &_
										"WHERE k_value = " & FormatDatabaseString(strKeywords, 255) & " " &_
										"AND k_site = " & intSite

									Dim rsKeywords : Set rsKeywords = Server.CreateObject("ADODB.Recordset")

									If strDatabaseType = "MYSQL" Then
										rsKeywords.CursorLocation = 3
									End If

									rsKeywords.Open strSql, objConn, 1, 2, &H0001

									If rsKeywords.Eof Then
										rsKeywords.AddNew
										rsKeywords(1) = ProtectInsert(strKeywords, 255)
										rsKeywords(2) = intSite
										rsKeywords.Update
									End If
									intKeywords = rsKeywords("k_id")
									rsKeywords.Close : Set rsKeywords = Nothing
								End If
							End If
						End If
					
						rsReferrer.Addnew
						rsReferrer(1) = ProtectInsert(strReferrer, 255)
						rsReferrer(2) = intReferrerName
						rsReferrer(3) = intKeywords
						rsReferrer.Update
					End If
					intReferrer = rsReferrer(0)
					
					rsReferrer.Close : Set rsReferrer = Nothing

					If Request.Cookies("tosh")("rid") = "" Then
						Response.Cookies("tosh")("rid") = intReferrer
						Response.Cookies("tosh").Expires = DateAdd("d", 3650, strDateTime)
					End If

				End If
	
				rsDefinitions.Close : Set rsDefinitions = Nothing
				
				Dim rsLog : Set rsLog = Server.Createobject("ADODB.Recordset")
	
				strSql = "INSERT INTO " & strTablePrefix & "PageLog (pl_datetime, pl_pn_id, pl_r_id, pl_s_id) VALUES(" &_
					FormatDatabaseDate(strDateTime) & ", " &_
					intPage & ", " &_
					intReferrer & ", " &_
					intSessionID & ")"

				rsLog.Open strSql, objConn, 0, 2, &H0001
				Set rsLog = Nothing
			
			Else
	
				intUserAgent = CheckName(2, strUserAgent)
				Dim intRobot : intRobot = CheckName(6, strRobot)
			
				Dim rsRobot : Set rsRobot = Server.Createobject("ADODB.Recordset")
	
				strSql = "INSERT INTO " & strTablePrefix & "RobotLog (rl_datetime, rl_pn_id, rl_useragent, rl_robot, rl_ip) VALUES(" &_
					FormatDatabaseDate(strDateTime) & ", " &_
					intPage & ", " &_
					intUserAgent & ", " &_
					intRobot & ", " &_
					intIPNumber & ")"

				rsRobot.Open strSql, objConn, 0, 2, &H0001
			
			End If
			
		End If
	
	End Sub

	Private Function ConvertIPAddressToLong(strIPAddress)

		Dim strTemp : strTemp = strIPAddress
		Dim aryIP : aryIP = Split(strTemp, ".")
		Dim intNumber : intNumber = (Int(aryIP(0)) * 16777216) + (Int(aryIP(1)) * 65536) + (Int(aryIP(2)) * 256) + Int(aryIP(3))

		intNumber = intNumber - 2147483647
		
		ConvertIPAddressToLong = intNumber

	End Function

	Private Function ExtractPath(strScriptName)

		Dim strTemp : strTemp = Left(strScriptName, InStrRev(strScriptName, "/"))

		ExtractPath = strTemp

	End Function

	Private Function ExtractFileType(strScriptName)

		Dim strTemp
		If InstrRev(strScriptName, ".") > 0 And Right(strScriptName, 1) <> "/" Then
			strTemp = Mid(strScriptName, InStrRev(strScriptName, ".") + 1)
		Else
			strTemp = ""
		End If

		ExtractFileType = strTemp

	End Function

	Private Function GetCountry(intIPNumber)
	
		Dim strValue

		If Not IsNumeric(intIPNumber) Then
			strValue = ""
		Else
			strSql = "SELECT ic_code FROM " & strTablePrefix & "IPCountry " &_
				"WHERE " & intIPNumber & " BETWEEN ic_ipstart and ic_ipend"

			Dim rsCountry : Set rsCountry = Server.CreateObject("ADODB.Recordset")
			rsCountry.Open strSql, objConn, 1, 2, 1

			If Not rsCountry.Eof Then
				strValue = rsCountry(0)
			Else
				strValue = ""
			End If

			rsCountry.Close
			Set rsCountry = Nothing
		End If

		GetCountry = strValue
	
	End Function
	
	Private Function CheckName(intType, strName)

		Dim intValue

		If strName = "" Then
			intValue = 0
		Else
			strSql = "SELECT n_id, n_value, n_type FROM " & strTablePrefix & "Names WHERE n_value = " & FormatDatabaseString(strName, 255)

			Dim rsName : Set rsName = Server.CreateObject("ADODB.Recordset")
			
			If strDatabaseType = "MYSQL" Then
				rsName.CursorLocation = 3
			End If
			
			rsName.Open strSql, objConn, 1, 2, &H0001

			If rsName.Eof Then
				rsName.AddNew
				rsName("n_value")		= ProtectInsert(strName, 255)
				rsName("n_type")		= intType
				rsName.Update
			End If
			intValue = rsName("n_id")

			rsName.Close
			Set rsName = Nothing
			
		End If
		
		CheckName = intValue

	End Function

	Public Function ExtractPage(strReferrer)
	
		Dim strTemp : strTemp = strReferrer
	
		If InStr(strTemp, "?") Then
			strTemp = Mid(strTemp, 1, InStr(strTemp, "?") - 1)
		End If
		
		If Left(LCase(strTemp), 4) <> "http" Then
			strTemp = ""
		End If
		
		ExtractPage = strTemp
		
	End Function
	
	Public Function ExtractHost(strReferrer)
		
		Dim strTemp : strTemp = strReferrer
		
		strTemp = Replace(strTemp, "http://", "")
		strTemp = Replace(strTemp, "https://", "")

		If InStr(strTemp, "/") > 0 Then
			strTemp = Mid(strTemp, 1, InStr(strTemp, "/") - 1)
		End If
		
		ExtractHost = strTemp
	
	End Function
	
	Public Function ExtractDomain(strHost)
			
		Dim strDomain, strExtension
		
		Dim strTemp : strTemp = strHost
		
		If InStr(strTemp, ".") > 0 Then
		
			Dim strEnd : strEnd = Mid(strTemp, InStrRev(strTemp, "."))
		
			If InStr(".com.net.org.edu.gov.mil.int.aero.biz.coop.info.museum.name.pro", strEnd) > 0 Then
				strExtension = strEnd
			Else
				If Len(strEnd) = 3 And Not IsNumeric(Right(strEnd, 2)) Then 
					Dim strRemainder : strRemainder = Left(strTemp, Len(strTemp) - Len(strEnd))
					Dim strPart : strPart = Right(strRemainder, Len(strRemainder) - InStrRev(strRemainder, ".") + 1)
					Dim strGeneric : strGeneric = ".ac.com.co.edu.go.gv.gov.govt.int.ltd.mi.mil.net.or.org.plc"

					Select Case strEnd
					Case ".ca"
						strExtension = CheckExtension(".ab.bc.mb.nb.nf.ns.nt.nu.on.pe.qc.sk.yk", strPart, strEnd)
					Case Else
						strExtension = CheckExtension(strGeneric, strPart, strEnd)
					End Select
					
					If strExtension = "" Then
						strExtension = strEnd
					End If
					
				End If
			End If
			
		End If

		If strExtension <> "" Then
		
			Dim objSearch : Set objSearch	= New RegExp
			
			Dim strPattern : strPattern = "[\w|\-]+" & Replace(strExtension, ".", "\.") & "$"
			
			With objSearch
				.Pattern 	= strPattern
				.IgnoreCase = True
				.Global 	= False
			End With

			Dim objResults : Set objResults = objSearch.Execute(strTemp)

			If objResults.Count > 0 Then
				Dim colItem
				For Each colItem In objResults
					strDomain = colItem.Value
					Exit For
				Next
			End If
			
			Set objSearch = Nothing : Set objResults = Nothing
		Else
			strDomain = ""
		End If

		ExtractDomain = strDomain
	
	End Function
	
	Private Function CheckExtension(strCompare, strPart, strEnd)
	
		Dim strTemp
	
		If InStr(strCompare, strPart) > 0 Then
			strTemp = strPart & strEnd
		End If
		
		CheckExtension = strTemp
	
	End Function
	
	Public Function ExtractExtension(strDomain)
	
		Dim strTemp : strTemp = strDomain
		
		If strDomain <> "" Then
			strTemp = Mid(strTemp, InStr(strTemp, "."))
		Else
			strTemp = ""
		End If
	
		ExtractExtension = strTemp
	
	End Function
	
	Private Function CleanLanguage(strLanguage)
	
		Dim strTemp : strTemp = strLanguage
		
		If InStr(Left(strTemp, 2), "rl") > 0 Then
			strTemp = ""
		End If
		
		If strTemp <> "" Then
			If InStr(strTemp, ",") > 0 Then
				strTemp = Trim(Left(strTemp, InStr(strTemp, ",") - 1))
			Else
				strTemp = Trim(strTemp)
			End If
			If InStr(strTemp, ";") > 0 Then
				strTemp = Trim(Left(strTemp, InStr(strTemp, ";") - 1))
			End If
		End If
	
		CleanLanguage = strTemp
	
	End Function
	
	Private Function MatchDefinition(rsDefinition, strCompare, intType)
	
		Dim strMatch

		rsDefinition.Filter = "d_type = " & intType
		
		Do While Not rsDefinition.Eof
			Dim objSearch : Set objSearch = New RegExp
			With objSearch
				.Pattern 	= rsDefinition(2)
				.IgnoreCase = True
				.Global 	= False
			End With
			
			'On Error Resume Next
			
			If objSearch.Test(strCompare) = True Then

				strMatch = rsDefinition(1)
				
				If intType = 4 Then
					strExtra = rsDefinition(3)
				End If
				
				Exit Do
			End If
			
			'On Error Goto 0

			Set objSearch = Nothing
			rsDefinition.Movenext
			
		Loop
		
		MatchDefinition = strMatch
	
	End Function
	
	Private Function ExtractKeywords(strReferrer, strPrefix)
	
		Dim strPattern
	
		Dim strKeywords : strKeywords = ""
		
		Dim strTemp : strTemp = Right(strReferrer, Len(strReferrer) - InStr(strReferrer, "?") + 1)
		
		If InStr(strPrefix, "/") > 0 Then
			strPattern = strPrefix & "(.+)"
		Else
			strPattern = "[\?|&]" & strPrefix & "=([^&]+)"
		End If
		
		Dim objSearch : Set objSearch = New RegExp
		With objSearch
			.Pattern 	= strPattern
			.IgnoreCase = True
			.Global 	= False
		End With
		Dim objResults : Set objResults = objSearch.Execute(strTemp)
		
		If objResults.Count > 0 Then
			Dim objMatch : Set objMatch = objResults(0)
			strKeywords = objMatch.SubMatches(0)
		End If
		
		Set objMatch = Nothing : Set objSearch = Nothing : Set objResults = Nothing

		If InStr(strKeywords, "&") > 0 Then
			strKeywords = Left(strKeywords, InStr(strKeywords, "&") - 1)
		End If
		
		strKeywords = UrlDecode(strKeywords)

		ExtractKeywords = strKeywords
	
	End Function
	
	Private Function ExtractScriptName(strScriptName)

		Dim strTemp : strTemp = strScriptName

		Dim objSearch : Set objSearch	= New RegExp
		With objSearch
			.Pattern 	= "(http|https)://[\w|\-|\.]+"
			.IgnoreCase	= True
			.Global 	= False
		End With
		
		strTemp = objSearch.Replace(strTemp, "")

		If Instr(strTemp, "?") > 0 Then
			strTemp = Mid(strTemp, 1, Instr(strTemp, "?") - 1)
		End If

		Set objSearch = Nothing

		ExtractScriptName = strTemp

	End Function
	
	Private Function ExtractQuerystring(strScriptName)

		Dim strQuerystring

		Dim strTemp : strTemp = strScriptName
		
		If Instr(strTemp, "?") > 0 Then
			strQuerystring = Mid(strTemp, Instr(strTemp, "?") + 1)
		Else
			strQuerystring = ""
		End If

		ExtractQuerystring = strQuerystring

	End Function
	
	Public Function MatchIPAddress(strIPList)
	
		Dim intLoop
		
		Dim blnMatch : blnMatch = False
		Dim aryIPAddress : aryIPAddress = Split(Replace(strIPList, " ", ""), ",")
		Dim strIPAddress : strIPAddress = Request.Servervariables("REMOTE_ADDR")
		
		For intLoop = 0 To UBound(aryIPAddress)

			If Instr(aryIPAddress(intLoop), "*") Then

				Dim aryIPAddressList : aryIPAddressList = Split(aryIPAddress(intLoop), ".")
				Dim aryIPAddressSource : aryIPAddressSource = Split(strIPAddress, ".")
				
				If UBound(aryIPAddressList) = 3 And UBound(aryIPAddressSource) = 3 Then
					If aryIPAddressList(2) = "*" Then
						aryIPAddressList(2) = aryIPAddressSource(2)
					End If
					If aryIPAddressList(3) = "*" Then
						aryIPAddressList(3) = aryIPAddressSource(3)
					End If

					Dim strIPAddressCheck : strIPAddressCheck = aryIPAddressList(0) & "." & aryIPAddressList(1) & "." & aryIPAddressList(2) & "." &  aryIPAddressList(3)

					If strIPAddress = strIPAddressCheck Then
						blnMatch = True
						Exit For
					End If
				End If
			Else
				If strIPAddress = aryIPAddress(intLoop) Then
					blnMatch = True
					Exit For
				End If
			End If
		Next
		
		MatchIPAddress = blnMatch
	
	End Function

	Private Function URLDecode(strDecode)

		Dim strSource, strTemp, strResult, intPos

  		strDecode = Replace(strDecode, "%C3%A4", "ä")
  		strDecode = Replace(strDecode, "%C3%B6", "ö")
  		strDecode = Replace(strDecode, "%E5", "å")
  		strDecode = Replace(strDecode, "%E4", "Ä")
  		strDecode = Replace(strDecode, "%F6", "Ö")
  		strDecode = Replace(strDecode, "%C3%A5", "Å")
		strDecode = Replace(strDecode, "%C3%B8", "ø")

		strSource = Replace(strDecode, "+", " ")
		
		For intPos = 1 To Len(strSource)
			strTemp = Mid(strSource, intPos, 1)
			If strTemp = "%" Then
				If intPos + 2 <= Len(strSource) Then
					strResult = strResult & Chr(CInt("&H" & Mid(strSource, intPos + 1, 2)))
					intPos = intPos + 2
				End If
			Else
				strResult = strResult & strTemp
			End If
		Next

		URLDecode = strResult
	
	End Function
	
	Private Function IsPrivateIP(strIPAddress)
	
		Dim blnCheck : blnCheck = False
		
		If Left(strIPAddress, 3) = "10." Then
			blnCheck = True
		ElseIf strIPAddress = "127.0.0.1" Then
			blnCheck = True
		ElseIf Left(strIPAddress, 7) = "192.168" Then
			blnCheck = True
		ElseIf Left(strIPAddress, 4) = "172." Then
			Dim aryIP : aryIP = Split(strIPAddress, ".")
			If UBound(aryIP) = 3 Then
				If CInt(aryIP(1)) => 16 And CInt(aryIP(1)) =< 31 Then
					blnCheck = True
				End If
			End If
		End If
	
		IsPrivateIP = blnCheck
	
	End Function
	
	Private Function ProtectInsert(strValue, intLength)
	
		ProtectInsert = Left(strValue, intLength)
	
	End Function
	
	Private Function FormatIPAddress(strIPAddress)
	
		Dim strTemp, aryIPAddress
	
		If InStr(strIPAddress, ".") > 0 Then
			
			aryIPAddress = Split(strIPAddress, ".")
			
			If UBound(aryIPAddress) <> 3 Then
				Exit Function
			End If
			
			If Not IsNumeric(aryIPAddress(0)) Then
				Exit Function
			ElseIf Not IsNumeric(aryIPAddress(1)) Then
				Exit Function
			ElseIf Not IsNumeric(aryIPAddress(2)) Then
				Exit Function
			ElseIf Not IsNumeric(aryIPAddress(3)) Then
				Exit Function
			Else
				strTemp = strIPAddress
			End If
	
		End If
	
		FormatIPAddress = strTemp
	
	End Function
	
	Private Function FilterQuerystring(strQuerystring)
	
		Dim strTemp, blnMatch, intLoop, intFilter, aryVariable
	
		If strQuerystringFilter <> "" And strQuerystring <> "" Then
		
			Dim aryQuerystringFilter : aryQuerystringFilter = Split(strQuerystringFilter, ",")
			Dim aryQuerystring : aryQuerystring = Split(strQuerystring, "&")
			
			For intLoop = 0 To UBound(aryQuerystring)
				blnMatch = False
				If InStr(aryQuerystring(intLoop), "=") Then
					aryVariable = Split(aryQuerystring(intLoop), "=")
					For intFilter = 0 To UBound(aryQuerystringFilter)
						If LCase(aryVariable(0)) = LCase(aryQuerystringFilter(intFilter)) Then
							blnMatch = True
							Exit For
						Else
							blnMatch = False
						End If
					Next
				End If
				If blnMatch = False Then
					If strTemp = "" Then
						strTemp = strTemp & aryQuerystring(intLoop)
					Else
						strTemp = strTemp & "&" & aryQuerystring(intLoop)
					End If
				End If
			Next
		Else
			strTemp = strQuerystring
		End If
		
		FilterQuerystring = strTemp
	
	End Function

	Private Function IsIPAddress(strIPAddress)
	
		Dim blnTemp, aryIPAddress
	
		blnTemp = False
	
		If InStr(strIPAddress, ".") > 0 Then
		
			aryIPAddress = Split(strIPAddress, ".")
			
			If UBound(aryIPAddress) = 3 Then
			
				If IsNumeric(aryIPAddress(0)) And IsNumeric(aryIPAddress(1)) _
				And IsNumeric(aryIPAddress(2)) And IsNumeric(aryIPAddress(3)) Then
					blnTemp = True
				End If
				
			End If
			
		End If
	
		IsIPAddress = blnTemp
	
	End Function
	
End Class

%>