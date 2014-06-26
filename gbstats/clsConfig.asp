<%
Class MTConfig

	' DEFINE CLASS ONLY VARIABLES
	Private strSql, datStart, datEnd, aryDB, aryConfig
	Private strDatabaseType, strTablePrefix
	Private intSetting
	Private strSiteName, strShortDate

	Public Property Let Config(pConfig)
		
		aryConfig = pConfig
		strSiteName = aryConfig(intMTSiteName)
		strShortDate = aryConfig(intMTShortDateFormat)
		
	End Property
	
	Public Property Let Database(pDatabase)
		aryDB = pDatabase
		' ASSIGN CONFIGS
		strDatabaseType 	= aryDB(0)
		strTablePrefix 		= aryDB(5)
	End Property
	
	Public Property Let Connection(pConnection)
		objConn = pConnection
	End Property
	
	Public Property Let Setting(pSetting)
		intSetting = pSetting
	End Property
	
	Public Function SiteName()
		SiteName = strSiteName
	End Function
	
	Public Function Version()
		
		Version = "FJstats 1.0"
		
	End Function
	
	' *****************
	' * SETTINGS	  *
	' *****************
	
	Public Sub GenerateSettings(intSetting)
	
		Dim strResult, rsQuery, intRow, strClass, intLoop
		Dim strMode, strAction, intID, strError, strMsg, strGroup
		
		' DISPLAY SETTINGS
		Select Case intSetting
		
		Case 0 ' General
			
			Call DisplaySettingsHeader("General")
			
			With Response
				strSql = "SELECT c_name, c_value, c_group FROM " & strTablePrefix & "Config " &_
					"WHERE c_type IN (0,1) " &_
					"ORDER BY c_type ASC, c_order ASC"
					
				Set rsConfig = Server.CreateObject("ADODB.RecordSet")
				rsConfig.Open strSql, objConn, 3, 1, &H0000
			
				Do While Not rsConfig.Eof
					If rsConfig(2) <> strGroup Then
						.Write("<tr><th colspan=2 align=left>" & rsConfig(2) & "</th></tr>")
					End If
					strGroup = rsConfig(2)
					.Write("<tr>")
					.Write("<td>" & Replace(rsConfig(0), "_", " ") & ":</td>")
					If rsConfig(1) <> "" Then
						.Write("<td>" & rsConfig(1) & "</td>")
					Else
						.Write("<td>N/A</td>")
					End If
					.Write("</tr>")
					rsConfig.Movenext
					intLoop = intLoop + 1
				Loop
		
			End With
			
			rsConfig.Close : Set rsConfig = Nothing
			
			' COUNT PAGELOG RECORDS
			Dim intPageLog
			strSql = "SELECT COUNT(*) FROM " & strTablePrefix & "PageLog"
			Dim rsPageLog : Set rsPageLog = Server.CreateObject("ADODB.RecordSet")
			rsPageLog.Open strSql, objConn, 3, 1, &H0000
			
			If Not rsPageLog.Eof Then
				intPageLog = rsPageLog(0)
			Else
				intPageLog = 0
			End If
			rsPageLog.Close : Set rsPageLog = Nothing
			
			' COUNT ROBOTLOG RECORDS
			Dim intRobotLog
			strSql = "SELECT COUNT(*) FROM " & strTablePrefix & "RobotLog"
			Dim rsRobotLog : Set rsRobotLog = Server.CreateObject("ADODB.RecordSet")
			rsRobotLog.Open strSql, objConn, 3, 1, &H0000
			If Not rsRobotLog.Eof Then
				intRobotLog = rsRobotLog(0)
			Else
				intRobotLog = 0
			End If
			rsRobotLog.Close : Set rsRobotLog = Nothing
			
			' COUNT USERS
			Dim intUsers
			strSql = "SELECT COUNT(*) FROM " & strTablePrefix & "Users"
			Dim rsUsers : Set rsUsers = Server.CreateObject("ADODB.RecordSet")
			rsUsers.Open strSql, objConn, 3, 1, &H0000
			If Not rsUsers.Eof Then
				intUsers = rsUsers(0)
			Else
				intUsers = 0
			End If
			rsUsers.Close : Set rsUsers = Nothing
			
			With Response
				.Write("<tr><th colspan=2 align=left>Logging</th></tr>")
				.Write("<tr>")
				.Write("<td>Log Records:</td>")
				.Write("<td>" & intPageLog & "</td>")
				.Write("</tr>")
				If intPageLog > 0 Then
					.Write("<tr>")
					.Write("<td>Log Start:</td>")
					.Write("<td>" & GetLogDate("MIN", "Page") & "</td>")
					.Write("</tr>")
					.Write("<tr>")
					.Write("<td>Log End:</td>")
					.Write("<td>" & GetLogDate("MAX", "Page") & "</td>")
					.Write("</tr>")
				End If
				.Write("<tr>")
				.Write("<td>Robot Log Records:</td>")
				.Write("<td>" & intRobotLog & "</td>")
				.Write("</tr>")
				If intRobotLog > 0 Then
					.Write("<tr>")
					.Write("<td>Robot Log Start:</td>")
					.Write("<td>" & GetLogDate("MIN", "Robot") & "</td>")
					.Write("</tr>")
					.Write("<tr>")
					.Write("<td>Robot Log End:</td>")
					.Write("<td>" & GetLogDate("MAX", "Robot") & "</td>")
					.Write("</tr>")
				End If
				.Write("<tr><th colspan=2 align=left>Other</th></tr>")
				.Write("<tr>")
				.Write("<td>Users:</td>")
				.Write("<td>" & intUsers & "</td>")
				.Write("</tr>")
			End With
			
			Call DisplaySettingsFooter()
		
		Case 1 ' CONFIG
		
			Dim objItem, aryConfig, strValue, aryExtra, rsConfig, strName, strRegExp
			
			If Request.Form("action") = "UPDATE" Then
			
				ReDim aryConfig(3, Request.Form.Count - 4)
				
				Dim aryForm : aryForm = Split(Request.Form, "&")
				For intLoop = 0 To UBound(aryConfig, 2)
					Dim aryFormItem : aryFormItem = Split(aryForm(intLoop), "=")
					aryConfig(0, intLoop) = aryFormItem(0)
					aryConfig(1, intLoop) = UrlDecode(aryFormItem(1))
				Next
				
				strSql = "SELECT c_extra, c_group FROM " & strTablePrefix & "Config " &_
					"WHERE c_type = 2 ORDER BY c_order ASC"
				Dim rsExtra : Set rsExtra = Server.CreateObject("ADODB.RecordSet")
				rsExtra.Open strSql, objConn, 3, 1, &H0000
				Dim intCounter
				Do While Not rsExtra.Eof
					aryConfig(2, intCounter) = rsExtra(0)
					aryConfig(3, intCounter) = rsExtra(1)
					rsExtra.Movenext : intCounter = intCounter + 1
				Loop
				rsExtra.Close : Set rsExtra = Nothing
				
				For intLoop = 0 To UBound(aryConfig, 2)

					aryExtra = Split(aryConfig(2, intLoop), "||")
					Select Case aryExtra(0)
					Case "textarea"
						aryConfig(1, intLoop) = Replace(aryConfig(1, intLoop), vbcrlf, ",")
					Case "checkbox"
						If CInt(aryConfig(1, intLoop)) = 1 Then
							aryConfig(1, intLoop) = True
						Else
							aryConfig(1, intLoop) = False
						End If
					End Select
					
					strRegExp = aryExtra(4)
					
					If strRegExp <> "" Then
					
						Dim objCheck : Set objCheck = New RegExp
						With objCheck
							.Pattern = strRegExp
							.IgnoreCase = True
							If Not .Test(aryConfig(1, intLoop)) Then
								strError = strError & "<li>" & Replace(aryConfig(0, intLoop), "_", " ") & "</li>" & vbcrlf
							End If
						End With
						Set objCheck = Nothing
					End If
					
				Next
				
				If strError = "" Then
					
					Dim strConfigPath : strConfigPath = Request.Servervariables("Script_Name")
					strConfigPath = Left(strConfigPath, InStrRev(strConfigPath, "/") - 1)
					
					Dim objFSO : Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
					
					On Error Resume Next
					Dim objTS : Set objTS = objFSO.OpenTextFile(Server.MapPath(strConfigPath & "/config.asp"), 2)
					
					If Err.Number = 0 Then
						
						objTS.WriteLine("<" & Chr(37))
						For intLoop = 0 To UBound(aryConfig, 2)
							strName = Replace(aryConfig(0, intLoop), "_", "")
							objTS.WriteLine("Dim intMT" & strName & " : " & "intMT" & strName & " = " & intLoop)
						Next
						objTS.WriteLine()
						objTS.WriteLine("Dim aryMTConfig(" & UBound(aryConfig, 2) & ")")
						For intLoop = 0 To UBound(aryConfig, 2)
							strName = Replace(aryConfig(0, intLoop), "_", "")
							aryExtra = Split(aryConfig(2, intLoop), "||")
							If aryExtra(1) = "str" Then
								strValue = """" & aryConfig(1, intLoop) & """"
							Else
								strValue = aryConfig(1, intLoop)
							End If
							objTS.WriteLine("aryMTConfig(" & "intMT" & strName & ") = " & strValue)
						Next
						
						objTS.WriteLine(Chr(37) & ">")
						
						strResult = "<p class=error>Successfully updated configuration</p>"
						
					Else
					
						strResult = "<p class=error>Configuration update failed. Details:<br><br>" &_
							"Description: <strong>" & Err.Description & "</strong></p>"
					
					End If
					
					On Error Goto 0
					
					Set objTS = Nothing : Set objFSO = Nothing
					
					' UPDATE DATABASE CONFIG
					strSql = "SELECT c_name, c_value " &_
						"FROM " & strTablePrefix & "Config " &_
						"WHERE c_type = 2 ORDER BY c_order ASC"
						
					Set rsConfig = Server.CreateObject("ADODB.RecordSet")
					rsConfig.Open strSql, objConn, 1, 2, &H0000
					
					intLoop = 0
					Do While Not rsConfig.Eof
						rsConfig.Update Array(0,1), Array(aryConfig(0, intLoop), aryConfig(1, intLoop))
						rsConfig.Movenext : intLoop = intLoop + 1
					Loop
						
					rsConfig.Close : Set rsConfig = Nothing
				
				Else
				
					strResult = "<p class=error>There were errors with the following fields:</p>" &_
						"<ul>" & strError & "</ul>"
				
				End If
				
			Else
			
				' GET CONFIG FROM DATABASE
				strSql = "SELECT c_name, c_value, c_extra, c_group FROM " & strTablePrefix & "Config " &_
					"WHERE c_type = 2 " &_
					"ORDER BY c_order ASC"
				
				Set rsConfig = Server.CreateObject("ADODB.Recordset")
				rsConfig.Open strSql, objConn, 1, 2, &H0000
				aryConfig = rsConfig.GetRows()
				rsConfig.Close : Set rsConfig = Nothing
				
			End If
		
			Call DisplaySettingsHeader("Configuration")
			
			If strResult <> "" Then
				Response.Write("<tr><td colspan=2>" & strResult & "</td></tr>")
			End If
			
			With Response
				.Write("<form method=post>")
				For intLoop = 0 To UBound(aryConfig, 2)
					If strGroup <> aryConfig(3, intLoop) Then
						.Write("<tr valign=top><th colspan=2 align=left>" & aryConfig(3, intLoop) & "</th></td>")
					End If
					strGroup = aryConfig(3, intLoop)
					.Write("<tr valign=top><td>" & Replace(aryConfig(0, intLoop), "_", " ") & "</td>")
					.Write("<td>")
					aryExtra = Split(aryConfig(2, intLoop), "||")
					Select Case aryExtra(0)
					Case "text"
						.Write("<input type=text size=" & aryExtra(2) & " maxlength=" & aryExtra(3) & " name=" & aryConfig(0, intLoop))
						.Write(" value=""" & aryConfig(1, intLoop) & """>")
					Case "checkbox"
						.Write("<input type=checkbox class=checkbox name=" & aryConfig(0, intLoop))
						.Write(SetInputState(CBool(aryConfig(1, intLoop)), True, " checked") & " value=1>")
					Case "textarea"
						.Write("<textarea cols=" & aryExtra(2) & " rows=" & aryExtra(3) & " name=" & aryConfig(0, intLoop) & ">" & Replace(aryConfig(1, intLoop), ",", vbcrlf) & "</textarea>")
					Case "select"
						.Write("<select name=" & aryConfig(0, intLoop) & ">")
						Dim arySelectNames : arySelectNames = Split(aryExtra(2), ",")
						Dim arySelectValues : If aryExtra(3) = "" Then
							arySelectValues = arySelectNames
						Else
							arySelectValues = Split(aryExtra(3), ",")
						End If
						Dim intSelect : For intSelect = 0 To UBound(arySelectNames)
							.Write("<option value=""" & arySelectValues(intSelect) & """" &_ 
								SetInputState(aryConfig(1, intLoop), arySelectValues(intSelect), " selected") &_ 
								">" & arySelectNames(intSelect) & "</option>")
						Next
						.Write("</select>")
					End Select
					.Write("</td></tr>")
				Next
				.Write("<input type=hidden name=action value=UPDATE>")
				.Write("<tr><td colspan=2 align=center><input type=image src=""images/update_btn.gif""></td></tr>")
				.Write("</form>")
			End With
			
			Call DisplaySettingsFooter()
		
		Case 2 ' USERS
			
			Dim strUsername, strPassword, blnAdmin
			
			strUsername = Request.Form("username")
			strPassword	= Request.Form("password")
			blnAdmin 	= CInt(Request.Form("admin"))
			intID 		= Request("id")
			strMode		= Request.Querystring("m")
			strAction	= Request.Form("action")
			
			If strMode = "" Then
				strMode = "VIEW"
			ElseIf strMode = "DELETE" Then
				strMode = "VIEW"
				strAction = "DELETE"
			End If
			
			If blnAdmin = 1 Then 
				blnAdmin = True
			Else
				blnAdmin = False
			End If
			
			Select Case strAction
			
			Case "UPDATE", "ADD NEW"
				
				' DO SOME CHECKS
				If strAction = "ADD NEW" Then 
					If strUsername = "" Then
						strError = "<p class=error>Username cannot be blank.</p>"
					Else
						strSql = "SELECT u_id FROM " & strTablePrefix & "Users WHERE u_username = " & FormatDatabaseString(strUsername, 20)
						Dim rsCheck : Set rsCheck = Server.CreateObject("ADODB.Recordset")
						rsCheck.Open strSql, objConn, 3, 1, &H0000
						
						If Not rsCheck.Eof Then
							strError = "<p class=error>Username already exists.</p>"
						End If
						rsCheck.Close : Set rsCheck = Nothing
					End If					
				End If
				
				If strError = "" Then
					
					strSql = "SELECT u_id, u_username, u_password, u_admin FROM " & strTablePrefix & "Users " &_
						"WHERE u_id = " & intID
					
					Dim rsUpdate : Set rsUpdate = Server.CreateObject("ADODB.Recordset")
					rsUpdate.Open strSql, objConn, 1, 2, &H0000
					
					If rsUpdate.Eof Then
						rsUpdate.AddNew
					End If
					
					rsUpdate(1) = strUsername
					rsUpdate(2) = strPassword
					rsUpdate(3) = blnAdmin
					rsUpdate.Update
					
					rsUpdate.Close : Set rsUpdate = Nothing
					
					strMode = "VIEW"
					
				End If
			
			Case "DELETE"
			
				strSql = "DELETE FROM " & strTablePrefix & "Users WHERE u_id = " & intID
				Dim rsDelete : Set rsDelete = Server.CreateObject("ADODB.Recordset")
				rsDelete.Open strSql, objConn, 3, 1, &H0000
				Set rsDelete = Nothing
				
			End Select
			
			Call DisplaySettingsHeader("Users")
			
			Select Case strMode
			
			Case "EDIT"
				
				Dim strReadonly
				
				If intID = 0 Then
					strAction = "ADD NEW"
				Else
					strAction = "UPDATE"
					strReadonly = " readonly"
					
					strSql = "SELECT u_id, u_username, u_password, u_admin " &_
						"FROM " & strTablePrefix & "Users WHERE u_id = " & intID
					Set rsQuery = Server.CreateObject("ADODB.Recordset")
					rsQuery.Open strSql, objConn, 3, 1, &H0000
					
					If Not rsQuery.Eof Then
						strUsername = rsQuery(1)
						strPassword = rsQuery(2)
						blnAdmin 	= rsQuery(3)
					End If
					
					rsQuery.Close : Set rsQuery = Nothing
					
				End If
				
				If strError <> "" Then
					Response.Write("<tr><td colspan=2>" & strError & "</td></tr>")
				End If
				
				With Response
					.Write("<form method=post><input type=hidden name=id value=" & intID & ">")
					.Write("<tr><td><strong>Username:</strong> </td>")
					.Write("<td>")
					.Write("<input type=text size=20 maxlength=20 name=Username value=""" & strUsername & """" & strReadonly & ">")
					.Write("</td></tr>")
					.Write("<tr><td><strong>Password:</strong> </td>")
					.Write("<td>")
					.Write("<input type=password size=20 maxlength=20 name=Password value=""" & strPassword & """>")
					.Write("</td></tr>")
					.Write("<tr><td><strong>Admin?</strong> </td>")
					.Write("<td>")
					.Write("<input type=checkbox name=Admin value=1 " & SetInputState(blnAdmin, True, " checked") & " class=checkbox>")
					.Write("</td></tr>")
					.Write("<input type=hidden name=action value=""" & strAction & """>")
					.Write("<tr><td colspan=2 align=center><input type=image name=action src=""images/" & Replace(strAction, " ", "_") & "_btn.gif""></td></tr>")
					.Write("</form>")
				End With
			
			Case "VIEW"
			
				strSql = "SELECT u_id, u_username, u_password, u_admin " &_
					"FROM " & strTablePrefix & "Users " &_
					"ORDER BY u_username ASC"
					
				Set rsQuery = Server.CreateObject("ADODB.Recordset")
				
				If strDatabaseType = "MYSQL" Then
					rsQuery.CursorLocation = 3
				End If
				
				rsQuery.Open strSql, objConn, 3, 1, &H0000
				intRow = 0
				With Response
					If rsQuery.Recordcount > 0 Then
						.Write("<tr><th>Username</th><th>Admin</th><th>Action</th></tr>")
						Do While Not rsQuery.Eof
						
							intRow = intRow + 1
							If (intRow Mod 2) = 1 Then
								strClass = "data"
							Else
								strClass = "dataalt"
							End If
							
							.Write("<tr class=" & strClass & ">")
							.Write("<td>" & rsQuery(1) & "</td>")
							.Write("<td align=center>" & DisplayAdmin(rsQuery(3)) & "</td>")
							.Write("<td><input type=image src=""images/edit_btn.gif"" onclick=""document.location='?s=2&m=EDIT&id=" & rsQuery(0) & "'"">&nbsp;")
							.Write("<input type=image src=""images/delete_btn.gif"" onclick=""document.location='?s=2&m=DELETE&id=" & rsQuery(0) & "'""></td>")
							.Write("</tr>")
							rsQuery.Movenext
						Loop
					Else
						.Write("<tr><td colspan=3>There are no users.</td></tr>")
					End If
					
					rsQuery.Close : Set rsQuery = Nothing
					
					.Write("<tr><td colspan=3 align=center>")
					.Write("<input type=image src=""images/add_new_btn.gif"" value=""ADD NEW"" onclick=""document.location='?s=2&m=EDIT&id=0'"">")
					.Write("</td></tr>")
				End With

			End Select
			
			Call DisplaySettingsFooter()
		
		Case 3 ' MAINTENANCE
			
			Call DisplaySettingsHeader("Maintenance")
			
			If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
			
				Dim strInstallPath : strInstallPath = Request.Servervariables("Script_Name")
				strInstallPath = Left(strInstallPath, InStrRev(strInstallPath, "/") - 1) & "/data"
	
				Dim objUpload : Set objUpload = New Upload
				
				On Error Resume Next
				objUpload.Save(Server.MapPath(strInstallPath))
				
				strAction = UCase(objUpload.Form("action"))
				
				If strAction = "DEFINITIONS" Or strAction = "COUNTRIES" Then
				
					If Err.Number <> 0 Then
						strMsg = "<p class=error>Upload failed: <ul><li>Could not write file. " &_
							"Check permissions and ensure the IUSR_machine account has MODIFY " &_
							"permissions to the data folder.</li></ul></p>"
					End If
					
					Dim strFileName
					
		    		Dim objFiles : objFiles = objUpload.UploadedFiles.Keys
					If (UBound(objFiles) <> -1) Then
			        	Dim objFile : For Each objFile in objUpload.UploadedFiles.Keys
			            	strFileName = objUpload.UploadedFiles(objFile).FileName
			         	Next
		    		End If
				
				End If
				
				On Error Goto 0
				
				Dim datStart : datStart = objUpload.Form("startdate")
				Dim datEnd : datEnd = objUpload.Form("enddate")
				
				If strMsg = "" Then
					' PROCESS ACTIONS
					Select Case strAction
					Case "DEFINITIONS"
						strMsg = UpdateDefinitions(strFileName)
					Case "COUNTRIES"
						strMsg = UpdateCountries(strFileName)
					Case "COMPACT"
						strMsg = CompactDatabase()
					Case "DELETE"
						strMsg = DeleteStatistics(datStart, datEnd)
					Case "ROBOTLOG"
						strMsg = DeleteRobotLog(datStart, datEnd)
					End Select
				End If
			
			Else
			
				' SET DEFAULTS
				If Not IsDate(datStart) Then
					datStart = GetLogDate("MIN", "Page")
					If Not IsDate(datStart) Then
						datStart = Date()
					Else
						datStart = FormatDateTime(datStart, 2)
					End If
				End If
				
				If Not IsDate(datEnd) Then
					datEnd = GetLogDate("MAX", "Page")
					If Not IsDate(datEnd) Then
						datEnd = Date()
					Else
						datEnd = FormatDateTime(datEnd, 2)
					End If
				End If
			
			End If
			
			With Response
				.Write("<form method=post name=maintenance enctype=""multipart/form-data"" action=""?s=3"" ")
				.Write("onsubmit=""return validateform(this,3);"">")
				If strMsg <> "" Then
					.Write("<tr><td>" & strMsg & "</td></tr>")
				End If
				.Write("<tr><td><input type=radio name=action value=definitions class=checkbox onclick=""showhideconfig('selectfile',1);showhideconfig('selectdate',0);"">&nbsp;")
				.Write("<strong>Update Definitions</strong><br>")
				.Write("Definitions allow FJstats to detect ")
				.Write("information about visitors on your site")
				.Write("</td></tr>")
				.Write("<tr><td><input type=radio name=action value=countries class=checkbox onclick=""showhideconfig('selectfile',1);showhideconfig('selectdate',0);"">&nbsp;")
				.Write("<strong>Update Country Data</strong><br>")
				.Write("Country data translates IP addresses to the originating country")
				.Write("</td></tr>")
				If strDatabaseType <> "MSACCESS" Then
					strClass = "display: none;"
				Else
					strClass = ""
				End If
				.Write("<tr style=""" & strClass & """><td><input type=radio name=action value=compact class=checkbox onclick=""showhideconfig('selectfile',0);showhideconfig('selectdate',0);"">&nbsp;")
				.Write("<strong>Compact / Repair Database</strong><br>")
				.Write("Optimize the database (MS ACCESS ONLY)")
				.Write("</td></tr>")
				.Write("<tr><td><input type=radio name=action value=delete class=checkbox onclick=""showhideconfig('selectfile',0);showhideconfig('selectdate',1);"">&nbsp;")
				.Write("<strong>Delete Statistics</strong><br>")
				.Write("Delete statistics data for a specified date range")
				.Write("</td></tr>")
				.Write("<tr><td><input type=radio name=action value=robotlog class=checkbox onclick=""showhideconfig('selectfile',0);showhideconfig('selectdate',1);"">&nbsp;")
				.Write("<strong>Delete Robot Statistics</strong><br>")
				.Write("Delete robot data for a specified date range")
				.Write("</td></tr>")
				.Write("<tr id=selectdate style=""display: none;"" class=retotal><td align=center><table border=0 cellpadding=2 cellspacing=0>")
				.Write("<tr><td colspan=2><p><strong>Select a date range:</strong></p></td></tr>")
				.Write("<tr><td><input type=text name=startdisplay value=""" & FormatDisplayDate(datStart, strShortDate) & """ size=12 readonly>&nbsp;")
				.Write("<input type=hidden name=startdate value=""" & datStart & """>")
				.Write("<a href=""javascript:calendar('maintenance.start',document.maintenance.startdate.value);"">")
				.Write("<img src=""images/calendar.gif"" border=0></a></td>")
				.Write("<td><input type=text name=enddisplay value=""" & FormatDisplayDate(datEnd, strShortDate) & """ size=12 readonly>&nbsp;")
				.Write("<input type=hidden name=enddate value=""" & datEnd & """>")
				.Write("<a href=""javascript:calendar('maintenance.end',document.maintenance.enddate.value);"">")
				.Write("<img src=""images/calendar.gif"" border=0></a></td></tr>")
				.Write("<tr><td><p>Start: </p></td><td><p>End: </p></td></tr>")
				.Write("</table></td></tr>")
				.Write("<tr id=selectfile style=""display: none;"" class=retotal><td align=center><table border=0 cellpadding=2 cellspacing=0>")
				.Write("<tr><td><p><strong>Upload file:</strong></p></td></tr>")
				If CheckDefaultDatabaseLocation = False Then
					.Write("<tr><td><input type=file name=file size=35></td></tr>")
				Else
					.Write("<tr><td style=""width:300px;""><p><small>File upload has been disabled because FJstats is configured to use ")
					.Write("MS Access with the default database location and file name.")
				End If
				.Write("</table></td></tr>")
				.Write("<tr><td colspan=2 align=center>")
				.Write("<input type=image src=""images/perform_maint_btn.gif"" name=submit value=""Perform Maintenance""></td></tr>")
				.Write("</form>")
			End With
			
			Call DisplaySettingsFooter()
		
		End Select 
	
	End Sub
	
	Private Function SetInputState(blnActual, blnRadio, strValue)
		
		Dim strTemp
		
		If blnActual = blnRadio Then
			strTemp = strValue
		End If
		
		SetInputState = strTemp
	
	End Function
	
	Private Function DisplayAdmin(blnValue)
		
		Dim strTemp
		
		blnValue = CBool(blnValue)
		
		If blnValue = True Then
			strTemp = "Yes"
		Else
			strTemp = "No"
		End If
		
		DisplayAdmin = strTemp
	
	End Function
	
	Private Sub DisplaySettingsHeader(strName)
	
		With Response
			.Write("<table border=0 cellpadding=0 cellspacing=0>")
			.Write("<tr><td><table border=0 cellpadding=0 cellspacing=0>")
			.Write("<tr><td width=22><img src=""images/white_arrow.gif""></td>")
			.Write("<td width=""100%""><span class=name>")
			.Write(strName & "</td><td align=right width=24>")
			.Write("<a href=""javascript: showhelp('settings','" & strName & "');""><img src=""images/help.gif"" alt=""Help"" border=0></a>")
			.Write("</td></tr></table></td></tr>")
			.Write("<tr><td colspan=3 class=settings>")
			.Write("<table border=0 cellpadding=4 cellspacing=0 width=""100%"" class=settings>")
		End With
	
	End Sub
	
	Private Sub DisplaySettingsFooter()
	
		With Response
			.Write ("</table></table>")
		End With
	
	End Sub
	
	' *****************
	' MISC DB FUNCTIONS
	' *****************
	
	Public Function GetLogDate(strType, strTable) ' MIN OR MAX
	
		Dim datTemp
		
		Dim strFirst : strFirst = LCase(Left(strTable, 1))
		
		strSql = "SELECT " & strType & "(" & strFirst & "l_datetime) FROM " & strTablePrefix & strTable & "Log"
		Dim rsDate : Set rsDate = Server.CreateObject("ADODB.Recordset")
		rsDate.Open strSql, objConn, 3, 1, &H0000
		
		If Not rsDate.Eof Then
			datTemp = rsDate(0)
		Else
			datTemp = Date()
		End If
		
		rsDate.Close : Set rsDate = Nothing
		
		GetLogDate = datTemp
	
	End Function
	
	' *********************
	' MAINTENANCE FUNCTIONS
	' *********************
	
	Private Function UpdateDefinitions(strFileName)
	
		Dim strError, strLine, aryLine, intLine, strResult
		
		Dim strInstallPath : strInstallPath = Request.Servervariables("Script_Name")
		strInstallPath = Left(strInstallPath, InStrRev(strInstallPath, "/") - 1) & "/data"

		Dim objFSO : Set objFSO = Server.CreateObject("Scripting.FileSystemObject")

		On Error Resume Next

		If strFileName = "" Then
			strFileName = "definitions.txt"
		End If
		
		Dim objTS : Set objTS = objFSO.OpenTextFile(Server.MapPath(strInstallPath) & "\" & strFileName, 1)
		
		strError = strError & CheckErrors(Err.Number, Err.Description)
		If strError <> "" Then
			UpdateDefinitions = "<p class=error>Update definitions failed:</p><ul>" & strError & "</ul>"
			Exit Function
		End If
		
		If strDatabaseType = "MSACCESS" Then
			strSql = "DELETE FROM " & strTablePrefix & "Definitions"
		Else
			strSql = "TRUNCATE TABLE " & strTablePrefix & "Definitions"
		End If
		
		Dim rsTruncate : Set rsTruncate = Server.CreateObject("ADODB.Recordset")
		rsTruncate.Open strSql, objConn, 1, 2, &H0000
		Set rsTruncate = Nothing
		
		strError = strError & CheckErrors(Err.Number, Err.Description)
		If strError <> "" Then
			UpdateDefinitions = "<p class=error>Update definitions failed:</p><ul>" & strError & "</ul>"
			Exit Function
		End If
		
		On Error Goto 0
		
		If Not objTS.AtEndOfStream Then
		
			' GET FILE HEADER
			Dim strFirstLine : strFirstLine = objTS.ReadLine
			
			' CHECK HEADER
			If Not CheckFileHeader("Definitions", strFirstLine) Then
				UpdateDefinitions = "<p class=error>The definitions file has does not have the correct header. Check the file and try again.</p>"
				Exit Function
			End If
			
			Dim datSerial : datSerial = Right(strFirstLine, 10)
			
			Dim rsInsert : Set rsInsert = Server.CreateObject("ADODB.Recordset")
			
			Do While Not objTS.AtEndOfStream
	
				intLine = objTS.Line
				strLine = objTS.Readline
				aryLine = Split(strLine,"||")
				If UBound(aryLine) = 4 Then
					If strDatabaseType = "MYSQL" Then
						aryLine(0) = Replace(aryLine(0), "\", "\\")
						aryLine(1) = Replace(aryLine(1), "\", "\\")
						aryLine(2) = Replace(aryLine(2), "\", "\\")
						aryLine(3) = Replace(aryLine(3), "\", "\\")
						aryLine(4) = Replace(aryLine(4), "\", "\\")
					End If
					strSql = "INSERT INTO " & strTablePrefix & "Definitions " &_
						"(d_name, d_regexp, d_extra, d_url, d_type) VALUES(" &_
						FormatDatabaseString(Trim(aryLine(0)), 255) & ", " &_
						FormatDatabaseString(Trim(aryLine(1)), 255) & ", " &_
						FormatDatabaseString(Trim(aryLine(2)), 255) & ", " &_
						FormatDatabaseString(Trim(aryLine(3)), 255) & ", " &_
						Trim(aryLine(4)) & ")"
					rsInsert.Open strSql, objConn, 1, 2, &H0000
				Else
					strError = strError & "<li>Error on line " & intLine & ".</li>"
				End If
			Loop
		
			Set rsInsert = Nothing : Set objFSO = Nothing
		
		End If
		
		If strError <> "" Then
			strResult = "<p class=error>Update definitions partially completed, some lines had errors:</p><ul>" & strError & "</ul>"
		Else
			strResult = "<p class=error>Definitions successfully updated.</p>"
			Call UpdateConfigValue("Definitions", datSerial)
		End If
		
		objTS.Close : Set objTS = Nothing
		
		UpdateDefinitions = strResult
		
	End Function
	
	Public Function UpdateCountries(strFileName)
	
		Dim strError, strLine, aryLine, intLine, strResult
		
		Dim strInstallPath : strInstallPath = Request.Servervariables("Script_Name")
		strInstallPath = Left(strInstallPath, InStrRev(strInstallPath, "/") - 1) & "/data"

		Dim objFSO : Set objFSO = Server.CreateObject("Scripting.FileSystemObject")

		If strFileName = "" Then
			strFileName = "countries.txt"
		End If
		
		On Error Resume Next

		Dim objTS : Set objTS = objFSO.OpenTextFile(Server.MapPath(strInstallPath) & "\" & strFileName, 1)
		
		strError = strError & CheckErrors(Err.Number, Err.Description)
		If strError <> "" Then
			UpdateCountries = "<p class=error>Update Countries failed:</p><ul>" & strError & "</ul>"
			Exit Function
		End If
		
		If strDatabaseType = "MSACCESS" Then
			strSql = "DELETE FROM " & strTablePrefix & "IPCountry"
		Else
			strSql = "TRUNCATE TABLE " & strTablePrefix & "IPCountry"
		End If
		
		Dim rsTruncate : Set rsTruncate = Server.CreateObject("ADODB.Recordset")
		rsTruncate.Open strSql, objConn, 1, 2, &H0000
		Set rsTruncate = Nothing
		
		strError = strError & CheckErrors(Err.Number, Err.Description)
		If strError <> "" Then
			UpdateCountries = "<p class=error>Update countries failed:</p><ul>" & strError & "</ul>"
			Exit Function
		End If
		
		On Error Goto 0
		
		If Not objTS.AtEndOfStream Then
		
			' GET FILE HEADER
			Dim strFirstLine : strFirstLine = objTS.ReadLine
			
			' CHECK HEADER
			If Not CheckFileHeader("Countries", strFirstLine) Then
				UpdateCountries = "<p class=error>The countries file has does not have the correct header. Check the file and try again.</p>"
				Exit Function
			End If
			
			Dim datSerial : datSerial = Right(strFirstLine, 10)
			
			Dim rsInsert : Set rsInsert = Server.CreateObject("ADODB.Recordset")
			
			Do While Not objTS.AtEndOfStream
				
				intLine = objTS.Line
				strLine = objTS.Readline
				aryLine = Split(strLine,"||")
				If UBound(aryLine) = 2 Then
					strSql = "INSERT INTO " & strTablePrefix & "IPCountry " &_
						"(ic_ipstart, ic_ipend, ic_code) VALUES(" &_
						Trim(aryLine(0)) & ", " &_
						Trim(aryLine(1)) & ", " &_
						FormatDatabaseString(Trim(aryLine(2)), 2) & ")"
					rsInsert.Open strSql, objConn, 1, 2, &H0000
				Else
					strError = strError & "<li>Error on line " & intLine & ".</li>"
				End If
			Loop

			Set rsInsert = Nothing : Set objFSO = Nothing
			
		End If
			
		If strError <> "" Then
			strResult = "<p class=error>Update countries partially completed, some lines had errors:</p><ul>" & strError & "</ul>"
		Else
			strResult = "<p class=error>Countries successfully updated.</p>"
			Call UpdateConfigValue("Countries", datSerial)
		End If
		
		objTS.Close : Set objTS = Nothing
		
		UpdateCountries = strResult
		
	End Function
	
	Private Function DeleteStatistics(datStart, datEnd)
	
		Dim strError, strResult
		
		On Error Resume Next
	
		strSql = "DELETE FROM " & strTablePrefix & "PageLog " &_
			"WHERE pl_datetime BETWEEN " & FormatDatabaseDate(datStart) &_
			" AND " & FormatDatabaseDate(datEnd & " 23:59:59")
			
		Dim rsDelete1 : Set rsDelete1 = Server.CreateObject("ADODB.Recordset")
		rsDelete1.Open strSql, objConn, 1, 2, &H0000
		Set rsDelete1 = Nothing
		
		strError = strError & CheckErrors(Err.Number, Err.Description)
	
		strSql = "DELETE FROM " & strTablePrefix & "PageNames " &_
			"WHERE pn_id NOT IN (SELECT DISTINCT pl_pn_id FROM " & strTablePrefix & "PageLog)"
			
		Dim rsDelete2 : Set rsDelete2 = Server.CreateObject("ADODB.Recordset")
		rsDelete2.Open strSql, objConn, 1, 2, &H0000
		Set rsDelete2 = Nothing

		strError = strError & CheckErrors(Err.Number, Err.Description)
		
		strSql = "DELETE FROM " & strTablePrefix & "Referrers " &_
			"WHERE r_id NOT IN (SELECT DISTINCT pl_r_id FROM " & strTablePrefix & "PageLog)"
		Dim rsDelete3 : Set rsDelete3 = Server.CreateObject("ADODB.Recordset")
		rsDelete3.Open strSql, objConn, 1, 2, &H0000
		Set rsDelete3 = Nothing
		
		strError = strError & CheckErrors(Err.Number, Err.Description)
		
		strSql = "DELETE FROM " & strTablePrefix & "ReferrerNames " &_
			"WHERE rn_id NOT IN (SELECT DISTINCT r_rn_id FROM " & strTablePrefix & "Referrers)"
			
		Dim rsDelete4 : Set rsDelete4 = Server.CreateObject("ADODB.Recordset")
		rsDelete4.Open strSql, objConn, 1, 2, &H0000
		Set rsDelete4 = Nothing
		
		strError = strError & CheckErrors(Err.Number, Err.Description)
		
		strSql = "DELETE FROM " & strTablePrefix & "Keywords " &_
			"WHERE k_id NOT IN (SELECT DISTINCT r_k_id FROM " & strTablePrefix & "Referrers)"
			
		Dim rsDelete5 : Set rsDelete5 = Server.CreateObject("ADODB.Recordset")
		rsDelete5.Open strSql, objConn, 1, 2, &H0000
		Set rsDelete5 = Nothing
		
		strError = strError & CheckErrors(Err.Number, Err.Description)
		
		strSql = "DELETE FROM " & strTablePrefix & "Sessions " &_
			"WHERE s_id NOT IN (SELECT DISTINCT pl_s_id FROM " & strTablePrefix & "PageLog)"
			
		Dim rsDelete6 : Set rsDelete6 = Server.CreateObject("ADODB.Recordset")
		rsDelete6.Open strSql, objConn, 1, 2, &H0000
		Set rsDelete6 = Nothing
		
		strError = strError & CheckErrors(Err.Number, Err.Description)
	
		strSql = "DELETE FROM " & strTablePrefix & "Names " &_
			"WHERE n_id NOT IN (SELECT DISTINCT s_hostname FROM " & strTablePrefix & "Sessions) " &_
			"AND n_id NOT IN (SELECT DISTINCT s_useragent FROM " & strTablePrefix & "Sessions) " &_
			"AND n_id NOT IN (SELECT DISTINCT s_browser FROM " & strTablePrefix & "Sessions) " &_
			"AND n_id NOT IN (SELECT DISTINCT s_os FROM " & strTablePrefix & "Sessions) " &_
			"AND n_id NOT IN (SELECT DISTINCT s_screenarea FROM " & strTablePrefix & "Sessions) " &_
			"AND n_id NOT IN (SELECT DISTINCT k_site FROM " & strTablePrefix & "Keywords)"
			
		Dim rsDelete7 : Set rsDelete7 = Server.CreateObject("ADODB.Recordset")
		rsDelete7.Open strSql, objConn, 1, 2, &H0000
		Set rsDelete7 = Nothing
		
		strError = strError & CheckErrors(Err.Number, Err.Description)

		On Error Goto 0
		
		If strError <> "" Then
			strResult = "<p class=error>The following database errors occured: </p><ul>" & strError & "</ul>"
		Else
			strResult = "<p class=error>Statistics deleted successfully.</p>"
		End If
		
		Call UpdateConfigValue("Delete_Log", Year(Date()) & "-" & FormatDatePart(Month(Date())) & "-" & FormatDatePart(Day(Date())))
		
		DeleteStatistics = strResult
	
	End Function
	
	Private Function DeleteRobotLog(datStart, datEnd)
	
		Dim strError, strResult
	
		On Error Resume Next
	
		strSql = "DELETE FROM " & strTablePrefix & "RobotLog " &_
			"WHERE rl_datetime BETWEEN " & FormatDatabaseDate(datStart) & " " &_
			"AND " & FormatDatabaseDate(datEnd & " 23:59:59")
			
		Dim rsDelete1 : Set rsDelete1 = Server.CreateObject("ADODB.Recordset")
		rsDelete1.Open strSql, objConn, 1, 2, &H0000
		Set rsDelete1 = Nothing
		
		strError = strError & CheckErrors(Err.Number, Err.Description)

		strSql = "DELETE FROM " & strTablePrefix & "PageNames " &_
			"WHERE pn_id NOT IN (SELECT DISTINCT rl_pn_id FROM " & strTablePrefix & "RobotLog)"
			
		Dim rsDelete2 : Set rsDelete2 = Server.CreateObject("ADODB.Recordset")
		rsDelete2.Open strSql, objConn, 1, 2, &H0000
		Set rsDelete2 = Nothing

		strError = strError & CheckErrors(Err.Number, Err.Description)
		
		strSql = "DELETE FROM " & strTablePrefix & "Names " &_
			"WHERE n_id NOT IN (SELECT DISTINCT rl_useragent FROM " & strTablePrefix & "RobotLog) " &_
			"AND n_id NOT IN (SELECT DISTINCT rl_robot FROM " & strTablePrefix & "RobotLog)"
			
		Dim rsDelete3 : Set rsDelete3 = Server.CreateObject("ADODB.Recordset")
		rsDelete3.Open strSql, objConn, 1, 2, &H0000
		Set rsDelete3 = Nothing
		
		strError = strError & CheckErrors(Err.Number, Err.Description)
		
		On Error Goto 0
		
		If strError <> "" Then
			strResult = "<p class=error>The following database errors occured: </p><ul>" & strError & "</ul>"
		Else
			strResult = "<p class=error>Robot log deleted successfully.</p>"
		End If
		
		Call UpdateConfigValue("Delete_Robot_Log", Year(Date()) & "-" & FormatDatePart(Month(Date())) & "-" & FormatDatePart(Day(Date())))
		
		DeleteRobotLog = strResult
	
	End Function
	
	Private Function CheckErrors(intNumber, strDescription)
	
		Dim strError
	
		If intNumber <> 0 Then
			strError = strError & "<li>" & strDescription & "</li>"
			Err.Clear
		End If
		
		CheckErrors = strError
		
	End Function
	
	' ******************
	' * MISC MAINTENANCE
	' ******************
	
	Private Function CheckFileHeader(strFileType, strHeader)
	
		Dim blnResult
		
		If Left(strHeader, Len(strFileType) + 3) <> "##" & strFileType & ":" Or Not IsDate(Right(strHeader, 10)) Then
			blnResult = False
		Else
			blnResult = True
		End If
		
		CheckFileHeader = blnResult
	
	End Function
	
	Private Sub UpdateConfigValue(strName, strValue)
	
		strSql = "UPDATE " & strTablePrefix & "Config " &_
			"SET c_value = " & FormatDatabaseString(strValue, 255) & " " &_
			"WHERE c_name = " & FormatDatabaseString(strName, 255)
		
		Dim rsUpdate : Set rsUpdate = Server.CreateObject("ADODB.RecordSet")
		
		rsUpdate.Open strSql, objConn, 1, 2, &H0000

		Set rsUpdate = Nothing
	
	End Sub
	
	Private Function UrlDecode(strDecode)
	
		Dim strIn : strIn = strDecode
		Dim strOut : Dim intLoop
	 	
		Dim intPos : intPos = InStr(strIn, "+")
		Do While intPos
			Dim strLeft : Dim strRight
			If intPos > 1 Then 
				strLeft = Left(strIn, intPos - 1)
			End If
			If intPos < Len(strIn) Then 
				strRight = Mid(strIn, intPos + 1)
			End If
			strIn = strLeft & " " & strRight
			intPos = InStr(strIn, "+")
			intLoop = intLoop + 1
	 	Loop
		intPos = InStr(strIn, "%")
		Do While intPos
		   	If intPos > 1 Then 
				strOut = strOut & Left(strIn, intPos - 1)
			End If
		 	strOut = strOut & Chr(CInt("&H" & Mid(strIn, intPos + 1, 2)))
		   	If intPos > (Len(strIn) - 3) Then 
				strIn = ""
			Else
				strIn = Mid(strIn, intPos + 3)
			End If
			intPos = Instr(strIn, "%")
		Loop
		
	 	URLDecode = strOut & strIn
		
	End Function
	
	' ************
	' * DATABASE *
	' ************
	
	Private Function CompactDatabase()
	
		Dim strLocation, strName, strError
		Dim strConn, strConnBak, strLocationType, strDB, strTempDB, strResult
		
		strLocation = aryMTDB(1)
		strName		= aryMTDB(2)
		
		' CREATE RANDOM NUMBER
		Dim intSeconds : intSeconds	= Second(Now())
		Dim intMinutes :intMinutes = Minute(Now())
		Dim intRandom : intRandom = intMinutes * intSeconds 

		Dim objFSO : Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
		
		' CHECK TO SEE IF THERE IS A COLON IN strLOCATION
		If Len(strLocation) > 2 Then
			If Mid(strLocation, 2, 1) = ":" Or Mid(strLocation, 1, 2) = "\\" Then
				' PATH USES A DRIVE LETTER, MUST BE ABSOLUTE
				strLocationType = "ABSOLUTE"
			Else
				strLocationType = "VIRTUAL"
			End If
		Else
			strLocationType = "VIRTUAL"
		End If
		
		If strLocationType = "ABSOLUTE" Then
			strDB		= strLocation & "\" & strName
			strTempDB	= strLocation & "\" & "db" & intRandom & ".mdb"
		Else ' VIRTUAL
			strDB		= Server.MapPath(strLocation & "/" & strName)
			strTempDB	= Server.MapPath(strLocation & "/" & "db" & intRandom & ".mdb")
		End If
	
		strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDB
		strConnBak = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strTempDB
		
		If IsObject(objConn) Then
			objConn.Close : Set objConn = Nothing
		End If
		
		Dim objJRO : Set objJRO = Server.CreateObject("JRO.JetEngine")
	 	objJRO.CompactDatabase strConn, strConnBak
		Set objJRO = Nothing
	
		On Error Resume Next
	
		If objFSO.FileExists(strDB) And objFSO.FileExists(strTempDB) Then
		
			objFSO.DeleteFile(strDB)
			strError = strError & CheckErrors(Err.Number, Err.Description)
			
			objFSO.MoveFile strTempDB, strDB
			strError = strError & CheckErrors(Err.Number, Err.Description)
			
			If strError <> "" Then
				strResult = "<p class=error>Compact / Repair failed:</p><ul>" & strError & "</ul>"
			Else
				strResult = "<p class=error>Compact / Repair completed successfully.</p>"
			End If
			
		Else
			strResult = "<p class=error>Compact and Repair failed.</p>"
		End If

		On Error Goto 0
		
		Set objFSO = Nothing
		
		Call CreateDatabaseConnection(1)
		Call UpdateConfigValue("Compact", Year(Date()) & "-" & FormatDatePart(Month(Date())) & "-" & FormatDatePart(Day(Date())))
		
		CompactDatabase = strResult
	
	End Function

	Private Function CheckDefaultDatabaseLocation()

		Dim blnCheck : blnCheck = False

		Dim strDBType		: strDBType = aryMTDB(0)
		Dim strDBLocation	: strDBLocation = aryMTDB(1)
		Dim strDBName		: strDBName = aryMTDB(2)

		If UCase(strDBType) = "MSACCESS" Then
				Dim strInstallPath : strInstallPath = Request.Servervariables("Script_Name")
				strInstallPath = Left(strInstallPath, InStrRev(strInstallPath, "/") - 1)

				If LCase(strInstallPath) = LCase(strDBLocation) And LCase(strDBName) = "db.mdb" Then
					blnCheck = True
				End If
		End If

		CheckDefaultDatabaseLocation = blnCheck

	End Function

End Class
%>