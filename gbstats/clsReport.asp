<%

Class MTReport

	' DEFINE CLASS ONLY VARIABLES
	Private strSql, datStart, datEnd, aryDB, aryConfig, intReport
	Private intItems, blnItems, strForm
	Private strDatabaseType, strTablePrefix, intSetting
	Private strSiteName, strSiteUrl, blnSiteAliases, blnTruncateUrl
	Private strSiteAliases, intSessionDuration, blnShowGraph, blnAdmin
	Private strShortDate, strLongDate, intTimeOffset
	
	' REQUIRED PROPERTIES FOR ALL METHODS
	Public Property Let Database(pDatabase)
		aryDB = pDatabase
		' ASSIGN CONFIGS
		strTablePrefix = aryDB(5)
		strDatabaseType = aryDB(0)
	End Property

	Public Property Let Config(pConfig)
		
		aryConfig = pConfig
		
		'ASSIGN CONFIG VALUES
		strSiteName 			= aryConfig(intMTSiteName)
		strSiteUrl 				= aryConfig(intMTSiteUrl)
		strSiteAliases 			= aryConfig(intMTSiteAliases)
		intSessionDuration 		= aryConfig(intMTSessionDuration)
		blnShowGraph 			= CBool(aryConfig(intMTShowGraph))
		blnTruncateUrl			= CBool(aryConfig(intMTTruncateUrls))
		strShortDate			= aryConfig(intMTShortDateFormat)
		strLongDate				= aryConfig(intMTLongDateFormat)
		intTimeOffset			= aryConfig(intMTTimeOffset)
		
		' EXTRA
		If strSiteAliases <> "" Then
			blnSiteAliases = True
		Else
			blnSiteAliases = False
		End If
		
		If IsNumeric(intSessionDuration) = False Then
			intSessionDuration = 60
		End If
		
	End Property
	
	' OTHER PROPERTIES
	Public Property Let Report(pReport)
		intReport = pReport
	End Property
	
	Public Property Let StartDate(pStartDate)
		If IsDate(pStartDate) = False Then
			datStart = FormatDateTime(DateAdd("h", intTimeOffset, Now()), 2)
		Else
			datStart = pStartDate
		End If
	End Property
	
	Public Property Let EndDate(pEndDate)
		If IsDate(pEndDate) = False Then
			datEnd = FormatDateTime(DateAdd("h", intTimeOffset, Now()), 2)
		Else
			datEnd = pEndDate
		End If
	End Property
	
	Public Property Let Items(pItems)
		If pItems <> "" Then
			intItems = CInt(pItems)
		Else
			intItems = 100
		End If
		If intItems = 0 Then
			blnItems = False
		Else
			blnItems = True
		End If
	End Property
	
	Public Function SiteName()
		SiteName = strSiteName
	End Function
	
	Public Function Version()
		
		Call CheckVersion("")
		Version = ""
		
	End Function
	
	Public Sub GenerateReport()
	
		Dim strDesc, strClass, intRow, rsQuery
		Dim intTotal, sngPercent, intMaxNumber
		Dim intDayLoop, datCurrent, intCount
		Dim blnMovenext, intHourLoop, strUrl
		Dim strRegistry
		
		If datStart = datEnd Then
			strDesc = FormatDisplayDate(datStart, strLongDate)
		Else
			strDesc = FormatDisplayDate(datStart, strLongDate) & " - " & FormatDisplayDate(datEnd, strLongDate)
		End If
		
		Dim aryReport : aryReport = GetReportArray()
		Dim strReportName : strReportName = aryReport(intReport, 0)
		Dim strReportGroup : strReportGroup = aryReport(intReport, 1)
		
		Select Case intReport
		
		Case 0 ' GENERAL (SUMMARY)
		
			Dim intOnline : intOnline = CountUsersOnline()
			Dim intPageViews : intPageViews = CountPageViews("", datStart, datEnd)
			Dim intDistinctPages : intDistinctPages = CountDistinctPages()
			Dim intDistinctUrls : intDistinctUrls = CountDistinctUrls()
			Dim intVisits : intVisits = CountVisits(datStart, datEnd)
			Dim intVisitors : intVisitors = CountVisitors(datStart, datEnd)
			Dim intSearches : intSearches = CountSearches(datStart, datEnd)

			Dim intReportHours : intReportHours	= DateDiff("n", datStart, datEnd & " 23:59:59") / 60
			Dim sngReportDays : sngReportDays = FormatNumber(intReportHours / 24, 2)
			
			Dim sngPagesPerVisit
			If intVisits > 0 Then
				sngPagesPerVisit = FormatNumber(intPageViews / intVisits, 2)
			Else
				sngPagesPerVisit = 0
			End If
	
			Dim sngPageViewsPerHour
			If intReportHours > 0 Then
				sngPageViewsPerHour	= FormatNumber(intPageViews / intReportHours, 2)
			Else
				sngPageViewsPerHour	= 0
			End If
	
			Dim sngPageViewsPerDay, sngVisitsPerDay, sngVisitorsPerDay
			If sngReportDays > 0 Then
				sngPageViewsPerDay	= FormatNumber(intPageViews / sngReportDays, 2)
				sngVisitsPerDay		= FormatNumber(intVisits / sngReportDays, 2)
				sngVisitorsPerDay	= FormatNumber(intVisitors / sngReportDays, 2)
			Else
				sngPageViewsPerDay	= 0
				sngVisitsPerDay		= 0
				sngVisitorsPerDay	= 0
			End If
		
			Call DisplayReportHeader(strReportGroup, strReportName, strDesc)
			
			If intDistinctPages > 0 Then
				With Response
					.Write("<table border=0 cellpadding=4 cellspacing=0 width=""100%"" class=dataalt>")
					.Write("<input type=hidden name=cols value=2>")
					.Write("<tr><td>Page Views: </td>")
					.Write("<td align=right><strong>" & intPageViews & "</strong></td></tr>")
					.Write("<input type=hidden name=data1 value=""Page Views: "">")
					.Write("<input type=hidden name=data2 value=""" & intPageViews & """>")
					.Write("<tr><td>Visits: </td>")
					.Write("<td align=right><strong>" & intVisits & "</strong></td></tr>")
					.Write("<input type=hidden name=data1 value=""Visits: "">")
					.Write("<input type=hidden name=data2 value=""" & intVisits & """>")
					.Write("<tr><td>Unique Visitors: </td>")
					.Write("<td align=right><strong>" & intVisitors & "</strong></td></tr>")
					.Write("<input type=hidden name=data1 value=""Unique Visitors: "">")
					.Write("<input type=hidden name=data2 value=""" & intVisitors & """>")
					.Write("<tr><td>Users Online: </td>")
					.Write("<td align=right><strong>" & intOnline & "</strong></td></tr>")
					.Write("<input type=hidden name=data1 value=""Users Online: "">")
					.Write("<input type=hidden name=data2 value=""" & intOnline & """>")
					.Write("<tr><td>Searches: </td>")
					.Write("<td align=right><strong>" & intSearches & "</strong></td></tr>")
					.Write("<input type=hidden name=data1 value=""Searches: "">")
					.Write("<input type=hidden name=data2 value=""" & intSearches & """>")
					.Write("<tr><td>Avg. Page Views Per Visit: </td>")
					.Write("<td align=right><strong>" & sngPagesPerVisit & "</strong></td></tr>")
					.Write("<input type=hidden name=data1 value=""Avg. Page Views Per Visit: "">")
					.Write("<input type=hidden name=data2 value=""" & sngPagesPerVisit & """>")
					.Write("<tr><td>Avg. Page Views Per Hour: </td>")
					.Write("<td align=right><strong>" & sngPageViewsPerHour & "</strong></td></tr>")
					.Write("<input type=hidden name=data1 value=""Avg. Page Views Per Hour: "">")
					.Write("<input type=hidden name=data2 value=""" & sngPageViewsPerHour & """>")
					.Write("<tr><td>Avg. Page Views Per Day: </td>")
					.Write("<td align=right><strong>" & sngPageViewsPerDay & "</strong></td></tr>")
					.Write("<input type=hidden name=data1 value=""Avg. Page Views Per Day: "">")
					.Write("<input type=hidden name=data2 value=""" & sngPageViewsPerDay & """>")
					.Write("<tr><td>Avg. Visits Per Day: </td>")
					.Write("<td align=right><strong>" & sngVisitsPerDay & "</strong></td></tr>")
					.Write("<input type=hidden name=data1 value=""Avg. Visits Per Day:  "">")
					.Write("<input type=hidden name=data2 value=""" & sngVisitsPerDay & """>")
					.Write("<tr><td>Avg. Unique Visitors Per Day: </td>")
					.Write("<td align=right><strong>" & sngVisitorsPerDay & "</strong></td></tr>")
					.Write("<input type=hidden name=data1 value=""Avg. Unique Visitors Per Day: "">")
					.Write("<input type=hidden name=data2 value=""" & sngVisitorsPerDay & """>")
					.Write("<tr><td>Unique Pages Logged: </td>")
					.Write("<td align=right><strong>" & intDistinctPages & "</strong></td></tr>")
					.Write("<input type=hidden name=data1 value=""Unique Pages Logged: "">")
					.Write("<input type=hidden name=data2 value=""" & intDistinctPages & """>")
					.Write("<tr><td>Unique URLs Logged: </td>")
					.Write("<td align=right><strong>" & intDistinctUrls & "</strong></td></tr>")
					.Write("<input type=hidden name=data1 value=""Unique URLs Logged: "">")
					.Write("<input type=hidden name=data2 value=""" & intDistinctUrls & """>")
					.Write("</table>")
				End With
			Else
				Response.Write("<p class=nodata>The database is empty.</p>")
			End If
			
			Call DisplayReportFooter()
			
		Case 1 ' WHO'S ONLINE (SUMMARY)
		
			Dim datNow : datNow	= DateAdd("h", intTimeOffset, Now())
				
			strSql = "SELECT s_ip, n_value AS s_hostname, " &_
				"s_country, MAX(pl_datetime) AS dc_lasthit, COUNT(pl_pn_id) " &_	
				"FROM (SELECT pl_datetime, pl_pn_id, s_ip, s_country, s_hostname " &_
				"FROM " & strTablePrefix & "PageLog, " & strTablePrefix & "Sessions " &_
				"WHERE pl_s_id = s_id " &_
				"AND pl_datetime BETWEEN " & FormatDatabaseDate(DateAdd("n", 0 - intSessionDuration, datNow)) &_
				" AND " & FormatDatabaseDate(datNow) & ") dt_PageLog"
				
			If strDatabaseType = "MYSQL" Then
				strSql = strSql & " LEFT JOIN " & strTablePrefix & "Names ON s_hostname = n_id " &_
					"GROUP BY s_ip, s_hostname, s_country " &_
					"ORDER BY dc_lasthit DESC"
			Else
				strSql = strSql & " LEFT JOIN " & strTablePrefix & "Names " &_
					"ON dt_pagelog.s_hostname = " & strTablePrefix & "Names.n_id " &_
					"GROUP BY s_ip, n_value, s_country " &_
					"ORDER BY MAX(pl_datetime) DESC"
			End If

			Set rsQuery = Server.CreateObject("ADODB.Recordset")
			
			If strDatabaseType = "MYSQL" Then
				rsQuery.CursorLocation = 3
			End If
			
			rsQuery.Open strSql, objConn, 3, 1, &H0000
			
			Call DisplayReportHeader(strReportGroup, strReportName, rsQuery.RecordCount & " User(s) Online")
			
			If Not rsQuery.Eof Then
				Response.Write("<table border=0 cellpadding=4 cellspacing=0 width=""100%"">") & vbcrlf
				Response.Write("<tr><th align=left>Host</th><th align=left>Country</th><th align=right>Last Access</th><th align=right>Page Views</th></tr>") & vbcrlf
				Response.Write("<input type=hidden name=data1 value=""Host"">")
				Response.Write("<input type=hidden name=data2 value=""Country"">")
				Response.Write("<input type=hidden name=data3 value=""Last Access"">")
				Response.Write("<input type=hidden name=data4 value=""Page Views"">")
				Response.Write("<input type=hidden name=cols value=4>")
			Else
				Response.Write("<p class=nodata>There are currently no active users.</p>") & vbcrlf
			End if
			Do While Not rsQuery.Eof
			
				intRow = intRow + 1
				If (intRow Mod 2) = 1 Then
					strClass = "data"
				Else
					strClass = "dataalt"
				End If
				
				strRegistry = ChooseIPWhois(rsQuery(2))
				
				With Response
					.Write("<tr class=" & strClass & ">") & vbcrlf
					.Write("<td align=left>") & vbcrlf
					If rsQuery(1) <> "" Then
						.Write(rsQuery(1) & "<br>")
						.Write(FormatIPAddress(ConvertIPNumberToAddress(rsQuery(0)), strRegistry, strClass) & vbcrlf)
						.Write("<input type=hidden name=data1 value=""" & rsQuery(1) & " - " & ConvertIPNumberToAddress(rsQuery(0)) & """>")
					Else
						.Write(FormatIPAddress(ConvertIPNumberToAddress(rsQuery(0)), strRegistry, strClass) & vbcrlf)
						.Write("<input type=hidden name=data1 value=""" & ConvertIPNumberToAddress(rsQuery(0)) & """>")
					End If
					.Write("</td>")
					.Write("<td align=left>" & ConvertCountryCode(rsQuery(2)) & "</td>")
					.Write("<input type=hidden name=data2 value=""" & FormatExportData(ConvertCountryCode(rsQuery(2))) & """>")
					.Write("<td align=right>" & FormatDateTime(rsQuery(3), 3) & "</td>")
					.Write("<input type=hidden name=data3 value=""" & FormatExportData(FormatDateTime(rsQuery(3), 3)) & """>")
					.Write("<td align=right>" & rsQuery(4) & "</td>") & vbcrlf
					.Write("<input type=hidden name=data4 value=""" & rsQuery(4) & """>")
					.Write("</tr>") & vbcrlf
				End With
				
				rsQuery.Movenext
			Loop

			If rsQuery.RecordCount > 0  Then
				Response.Write("</table>") & vbcrlf
			End if
			
			rsQuery.Close : Set rsQuery = Nothing
			
			Call DisplayReportFooter()
			Call DisplayIPWhois()
		
		Case 2 : ' PAGES (PAGES & FILES)
				
			intTotal = CountPageViews("", datStart, datEnd)
			
			strSql = "SELECT " & SetItems(blnItems, intItems, True) & "pn_page, COUNT(pn_page) AS dc_pagecount FROM " &_
				"(SELECT pl_datetime, pn_page FROM " & strTablePrefix & "PageLog, " & strTablePrefix & "PageNames " &_
				"WHERE pl_pn_id = pn_id " &_
				"AND pl_datetime BETWEEN " & FormatDatabaseDate(datStart & " 00:00:00") &_
				" AND " & FormatDatabaseDate(datEnd & " 23:59:59") & ") dtPageLog " &_
				"GROUP BY pn_page "
			
			If strDatabaseType = "MYSQL" Then
				strSql = strSql & "ORDER BY dc_pagecount DESC, pn_page ASC " & SetItems(blnItems, intItems, False)
			Else
				strSql = strSql & "ORDER BY COUNT(pn_page) DESC, pn_page ASC"
			End If
			
			Set rsQuery = Server.CreateObject("ADODB.Recordset")

			If strDatabaseType = "MYSQL" Then
				rsQuery.CursorLocation = 3
			End If

			rsQuery.Open strSql, objConn, 3, 1, &H0000
			
			Call DisplayReportHeader(strReportGroup, strReportName, strDesc)
			
			If Not rsQuery.Eof Then
				Response.Write("<table border=0 cellpadding=4 cellspacing=0 width=""100%"">") & vbcrlf
				Response.Write("<tr><th></th><th align=left>Page</th><th>Count</th><th>%</th><th></th></tr>") & vbcrlf
				Response.Write("<input type=hidden name=data1 value=""Page"">")
				Response.Write("<input type=hidden name=data2 value=""Count"">")
				Response.Write("<input type=hidden name=data3 value=""%"">")
				Response.Write("<input type=hidden name=cols value=3>")
			Else
				Response.Write("<p class=nodata>There is no data available for the report period selected.</p>") & vbcrlf
			End if
			Do While Not rsQuery.Eof

				If intTotal > 0 Then
					sngPercent = FormatPercent(rsQuery(1) / intTotal, 2)
				Else
					sngPercent = FormatPercent(0, 2)
				End If
			
				intRow = intRow + 1
				If (intRow Mod 2) = 1 Then
					strClass = "data"
				Else
					strClass = "dataalt"
				End If
			
				If intRow = 1 Then
					intMaxNumber = rsQuery(1)
				End If
			
				With Response
					.Write("<tr class=" & strClass & ">") & vbcrlf
					.Write("<td align=right>" & intRow & ".</td>")
					.Write("<td align=left>" & FormatLink(rsQuery(0), strClass) & "</td>")
					.Write("<td align=right>" & rsQuery(1) & "</td>")
					.Write("<td align=right>" & sngPercent & "</td>")
					.Write("<td align=left>")
					If blnShowGraph = True Then
						Call GenerateGraph(rsQuery(1), intMaxNumber, strClass)
					End If
					.Write("</td>")
					.Write("</tr>") & vbcrlf
					.Write("<input type=hidden name=data1 value=""" & FormatExportData(rsQuery(0)) & """>")
					.Write("<input type=hidden name=data2 value=" & rsQuery(1) & ">")
					.Write("<input type=hidden name=data3 value=""" & sngPercent & """>") & vbcrlf
				End With
				rsQuery.Movenext
			Loop
			
			If rsQuery.RecordCount > 0  Then
				With Response
					.Write("<tr class=total>")
					.Write("<td colspan=2 align=right>Total: </td>")
					.Write("<td align=right>" & intTotal & "</td>")
					.Write("<td align=right>&nbsp;</td>")
					.Write("<td align=right>&nbsp;</td>")
					.Write("</tr>")
					.Write("<input type=hidden name=total value=" & intTotal & ">")
					.Write("</table>") & vbcrlf
				End With
			End if
			
			rsQuery.Close : Set rsQuery = Nothing
			
			Call DisplayReportFooter()
		
		Case 3 : ' URLS (PAGES & FILES)
			
			intTotal = CountPageViews("", datStart, datEnd)
			
			strSql = "SELECT " & SetItems(blnItems, intItems, True) & "pn_url, COUNT(pn_url) AS dc_pagecount, pn_label FROM " &_
				"(SELECT pn_url, pn_label FROM " & strTablePrefix & "PageLog, " & strTablePrefix & "PageNames " &_
				"WHERE pl_pn_id = pn_id " &_
				"AND pl_datetime BETWEEN " & FormatDatabaseDate(datStart & " 00:00:00") &_
				" AND " & FormatDatabaseDate(datEnd & " 23:59:59") & ") dtPageLog " &_
				"GROUP BY pn_url, pn_label "
				
			If strDatabaseType = "MYSQL" Then
				strSql = strSql & "ORDER BY dc_pagecount DESC, pn_url ASC " & SetItems(blnItems, intItems, False)
			Else
				strSql = strSql & "ORDER BY Count(pn_url) DESC, pn_url ASC"
			End If
			
			Set rsQuery = Server.CreateObject("ADODB.Recordset")
			
			If strDatabaseType = "MYSQL" Then
				rsQuery.CursorLocation = 3
			End If
			
			rsQuery.Open strSql, objConn, 3, 1, &H0000
			
			Call DisplayReportHeader(strReportGroup, strReportName, strDesc)
			
			If Not rsQuery.Eof Then
				Response.Write("<table border=0 cellpadding=4 cellspacing=0 width=""100%"">") & vbcrlf
				Response.Write("<tr><th></th><th align=left>Url</th><th>Count</th><th>%</th><th></th></tr>") & vbcrlf
				Response.Write("<input type=hidden name=data1 value=""Url"">")
				Response.Write("<input type=hidden name=data2 value=""Count"">")
				Response.Write("<input type=hidden name=data3 value=""%"">")
				Response.Write("<input type=hidden name=cols value=3>")
			Else
				Response.Write("<p class=nodata>There is no data available for the report period selected.</p>") & vbcrlf
			End if
			Do While Not rsQuery.Eof

				If intTotal > 0 Then
					sngPercent = FormatPercent(rsQuery(1) / intTotal, 2)
				Else
					sngPercent = FormatPercent(0, 2)
				End If
			
				intRow = intRow + 1
				If (intRow Mod 2) = 1 Then
					strClass = "data"
				Else
					strClass = "dataalt"
				End If
			
				If intRow = 1 Then
					intMaxNumber = rsQuery(1)
				End If
			
				With Response
					.Write("<tr class=" & strClass & ">") & vbcrlf
					.Write("<td align=right>" & intRow & ".</td>")
					.Write("<td align=left>")
					If rsQuery(2) <> "" Then
						.Write(rsQuery(2) & "<br>")
					End If
					.Write(FormatLink(rsQuery(0), strClass) & "</td>")
					.Write("<td align=right>" & rsQuery(1) & "</td>")
					.Write("<td align=right>" & sngPercent & "</td>")
					.Write("<td align=left>")
					If blnShowGraph = True Then
						Call GenerateGraph(rsQuery(1), intMaxNumber, strClass)
					End If
					.Write("</td>")
					.Write("</tr>") & vbcrlf
					If rsQuery(2) <> "" Then
						.Write("<input type=hidden name=data1 value=""" & FormatExportData(rsQuery(2) & " (" & rsQuery(0)) & ")"">")
					Else	
						.Write("<input type=hidden name=data1 value=""" & FormatExportData(rsQuery(0)) & """>")
					End If
					.Write("<input type=hidden name=data2 value=" & rsQuery(1) & ">")
					.Write("<input type=hidden name=data3 value=""" & sngPercent & """>") & vbcrlf
				End With
				rsQuery.Movenext
			Loop
			
			If rsQuery.RecordCount > 0  Then
				With Response
					.Write("<tr class=total>")
					.Write("<td colspan=2 align=right>Total: </td>")
					.Write("<td align=right>" & intTotal & "</td>")
					.Write("<td align=right>&nbsp;</td>")
					.Write("<td align=right>&nbsp;</td>")
					.Write("</tr>")
					.Write("<input type=hidden name=total value=" & intTotal & ">")
					.Write("</table>") & vbcrlf
				End With
			End if
			
			rsQuery.Close : Set rsQuery = Nothing
			
			Call DisplayReportFooter()
		
		Case 4 : ' DAILY (PAGES & FILES)
			
			intTotal 		= CountPageViews("", datStart, datEnd)
			If blnShowGraph = True Then
				intMaxNumber 	= GetMaxPageViews("DAILY")
			End If
			
			datCurrent		= datStart
			
			strSql	= "SELECT YEAR(pl_datetime), MONTH(pl_datetime), DAY(pl_datetime), COUNT(pl_pn_id) " &_
				"FROM " & strTablePrefix & "PageLog " &_
				"WHERE pl_datetime BETWEEN " & FormatDatabaseDate(datStart & " 00:00:00")&_
				" AND " & FormatDatabaseDate(datEnd & " 23:59:59") & " " &_
				"GROUP BY YEAR(pl_datetime), MONTH(pl_datetime), DAY(pl_datetime) " &_
				"ORDER BY YEAR(pl_datetime) ASC, MONTH(pl_datetime) ASC, DAY(pl_datetime) ASC"
			
			Set rsQuery = Server.CreateObject("ADODB.Recordset")
			rsQuery.Open strSql, objConn, 3, 1, &H0000
			
			Call DisplayReportHeader(strReportGroup, strReportName, strDesc)
			
			With Response
				.Write("<table border=0 cellpadding=4 cellspacing=0 width=""100%"">") & vbcrlf
				.Write("<tr><th align=left>Day</th><th align=right>Count</th><th>%</th><th></th></tr>") & vbcrlf
				.Write("<input type=hidden name=data1 value=""Day"">")
				.Write("<input type=hidden name=data2 value=""Count"">")
				.Write("<input type=hidden name=data3 value=""%"">")
			End With
			
			For intDayLoop = 1 to DateDiff("d", datStart, datEnd) + 1
			
				If Not rsQuery.Eof Then

					If DateDiff("d", datCurrent, DateSerial(rsQuery(0), rsQuery(1), rsQuery(2))) > 0 Then
						intCount = 0
						blnMovenext = False
					Else
						intCount = rsQuery(3)
						blnMovenext = True
					End If
				Else
					intCount = 0
					blnMovenext = False
				End If

				If intTotal > 0 Then
					sngPercent = FormatPercent(intCount / intTotal, 2)
				Else
					sngPercent = FormatPercent(0, 2)
				End If
			
				intRow = intRow + 1
				If (intRow Mod 2) = 1 Then
					strClass = "data"
				Else
					strClass = "dataalt"
				End If
			
				With Response
					.Write("<tr class=" & strClass & ">") & vbcrlf
					.Write("<td align=left>" & FormatDisplayDate(datCurrent, strShortDate) & "</td>")
					.Write("<td align=right>" & intCount & "</td>")
					.Write("<td align=right>" & sngPercent & "</td>")
					.Write("<td align=left>")
					If blnShowGraph = True Then
						Call GenerateGraph(intCount, intMaxNumber, strClass)
					End If
					.Write("</td>")
					.Write("</tr>") & vbcrlf
					.Write("<input type=hidden name=data1 value=""" & FormatDisplayDate(datCurrent, strShortDate) & """>")
					.Write("<input type=hidden name=data2 value=" & intCount & ">")
					.Write("<input type=hidden name=data3 value=""" & sngPercent & """>")
				End With
				
				datCurrent = DateAdd("d", 1, datCurrent)
				
				If blnMovenext = True Then
					rsQuery.Movenext
				End If
			Next
			
			With Response
					.Write("<tr class=total>")
					.Write("<td align=right>Total: </td>")
					.Write("<td align=right>" & intTotal & "</td>")
					.Write("<td align=right>&nbsp;</td>")
					.Write("<td align=right>&nbsp;</td>")
					.Write("</tr>")
					.Write("<input type=hidden name=total value=" & intTotal & ">")
					.Write("</table><br>") & vbcrlf
			End With
			
			rsQuery.Close : Set rsQuery = Nothing
			
			Call DisplayReportFooter()
		
		Case 5 : ' HOURLY (PAGES & FILES)
			
			intTotal 		= CountPageViews("", datStart, datEnd)
			If blnShowGraph = True Then
				intMaxNumber 	= GetMaxPageViews("HOURLY")
			End If
			
			If strDatabaseType = "MSSQL" Then
				strSql	= "SELECT DATEPART(hh, pl_datetime), COUNT(pl_pn_id) " &_
					"FROM " & strTablePrefix & "PageLog " &_
					"WHERE pl_datetime BETWEEN " & FormatDatabaseDate(datStart & " 00:00:00")&_
					" AND " & FormatDatabaseDate(datEnd & " 23:59:59") & " " &_
					"GROUP BY DATEPART(hh, pl_datetime) " &_
					"ORDER BY DATEPART(hh, pl_datetime) ASC"
			Else
				strSql	= "SELECT HOUR(pl_datetime), COUNT(pl_pn_id) " &_
					"FROM " & strTablePrefix & "PageLog " &_
					"WHERE pl_datetime BETWEEN " & FormatDatabaseDate(datStart & " 00:00:00")&_
					" AND " & FormatDatabaseDate(datEnd & " 23:59:59") & " " &_
					"GROUP BY HOUR(pl_datetime) " &_
					"ORDER BY HOUR(pl_datetime) ASC"
			End If
				
			Set rsQuery = Server.CreateObject("ADODB.Recordset")
			rsQuery.Open strSql, objConn, 3, 1, &H0000
			
			Call DisplayReportHeader(strReportGroup, strReportName, strDesc)
			
			With Response
				.Write("<table border=0 cellpadding=4 cellspacing=0 width=""100%"">") & vbcrlf
				.Write("<tr><th align=left>Hour</th><th align=right>Count</th><th>%</th><th></th></tr>") & vbcrlf
				.Write("<input type=hidden name=data1 value=""Hour"">")
				.Write("<input type=hidden name=data2 value=""Count"">")
				.Write("<input type=hidden name=data3 value=""%"">")
			End With
			
			For intHourLoop = 0 To 23
				
				If Not rsQuery.Eof Then
					If rsQuery(0) > intHourLoop Then
						intCount = 0
						blnMovenext = False
					Else
						intCount = rsQuery(1)
						blnMovenext = True
					End If
				Else
					intCount = 0
					blnMovenext = False
				End If

				If intTotal > 0 Then
					sngPercent = FormatPercent(intCount / intTotal, 2)
				Else
					sngPercent = FormatPercent(0, 2)
				End If
			
				intRow = intRow + 1
				If (intRow Mod 2) = 1 Then
					strClass = "data"
				Else
					strClass = "dataalt"
				End If
				
				With Response
					.Write("<tr class=" & strClass & ">")
					.Write("<td align=left>" & intHourLoop & ":00</td>")
					.Write("<td align=right>" & intCount & "</td>")
					.Write("<td align=right>" & sngPercent & "</td>")
					.Write("<td align=left>")
					If blnShowGraph = True Then
						Call GenerateGraph(intCount, intMaxNumber, strClass)
					End If
					.Write("</td>")
					.Write("</tr>")
					.Write("<input type=hidden name=data1 value=""" & intHourLoop & ":00"">")
					.Write("<input type=hidden name=data2 value=" & intCount & ">")
					.Write("<input type=hidden name=data3 value=""" & sngPercent & """>")
				End With
				
				If blnMovenext = True Then
					rsQuery.Movenext
				End If
			Next
			
			With Response
					.Write("<tr class=total>")
					.Write("<td align=right>Total: </td>")
					.Write("<td align=right>" & intTotal & "</td>")
					.Write("<td align=right>&nbsp;</td>")
					.Write("<td align=right>&nbsp;</td>")
					.Write("</tr>")
					.Write("<input type=hidden name=total value=" & intTotal & ">")
					.Write("</table><br>") & vbcrlf
			End With
			
			rsQuery.Close : Set rsQuery = Nothing
			
			Call DisplayReportFooter()
		
		Case 6 : ' BY IP ADDRESS (PAGES & FILES)
			
			intTotal = CountPageViews("", datStart, datEnd)

			strSql = "SELECT " & SetItems(blnItems, intItems, True) & "s_ip, n_value AS s_hostname, " &_
				"COUNT(pl_pn_id) AS dc_pagecount, s_country " &_
				"FROM (SELECT pl_datetime, pl_pn_id, s_ip, s_country, s_hostname " &_
				"FROM " & strTablePrefix & "PageLog, " & strTablePrefix & "Sessions " &_
				"WHERE pl_s_id = s_id " &_
				"AND pl_datetime BETWEEN " & FormatDatabaseDate(datStart & " 00:00:00") &_
				" AND " & FormatDatabaseDate(datEnd & " 23:59:59") & ") dt_PageLog"
				
			If strDatabaseType = "MYSQL" Then
				strSql = strSql & " LEFT JOIN " & strTablePrefix & "Names ON s_hostname = n_id " &_
					"GROUP BY s_ip, s_hostname, s_country " &_
					"ORDER BY dc_pagecount DESC, s_ip ASC " & SetItems(blnItems, intItems, False)
			Else
				strSql = strSql & " LEFT JOIN " & strTablePrefix & "Names " &_
					"ON dt_PageLog.s_hostname = " & strTablePrefix & "Names.n_id " &_
					"GROUP BY s_ip, n_value, s_country " &_
					"ORDER BY COUNT(pl_pn_id) DESC, s_ip ASC"
			End If

			Set rsQuery = Server.CreateObject("ADODB.Recordset")
			
			If strDatabaseType = "MYSQL" Then
				rsQuery.CursorLocation = 3
			End If
			
			rsQuery.Open strSql, objConn, 3, 1, &H0000
			
			Call DisplayReportHeader(strReportGroup, strReportName, strDesc)
			
			If Not rsQuery.Eof Then
				Response.Write("<table border=0 cellpadding=4 cellspacing=0 width=""100%"">") & vbcrlf
				Response.Write("<tr><th></th><th align=left>IP Address</th><th align=right>Page Views</th><th>%</th><th></th></tr>") & vbcrlf
				Response.Write("<input type=hidden name=data1 value=""IP Address"">")
				Response.Write("<input type=hidden name=data2 value=""Page Views"">")
				Response.Write("<input type=hidden name=data3 value=""%"">")
				Response.Write("<input type=hidden name=cols value=3>")
			Else
				Response.Write("<p class=nodata>There is no data available for the report period selected.</p>") & vbcrlf
			End if
			
			Do While Not rsQuery.Eof

				strRegistry = ChooseIPWhois(rsQuery(3))
			
				If intTotal > 0 Then
					sngPercent = FormatPercent(rsQuery(2) / intTotal, 2)
				Else
					sngPercent = FormatPercent(0, 2)
				End If
			
				intRow = intRow + 1
				If (intRow Mod 2) = 1 Then
					strClass = "data"
				Else
					strClass = "dataalt"
				End If
			
				If intRow = 1 Then
					intMaxNumber = rsQuery(2)
				End If
			
				With Response
					.Write("<tr class=" & strClass & ">") & vbcrlf
					.Write("<td align=right>" & intRow & ".</td>")
					.Write("<td align=left>") & vbcrlf
					If rsQuery(1) <> "" Then
						.Write(rsQuery(1) & "<br>")
						.Write(FormatIPAddress(ConvertIPNumberToAddress(rsQuery(0)), strRegistry, strClass) & vbcrlf)
						.Write("<input type=hidden name=data1 value=""" & rsQuery(1) & " - " & ConvertIPNumberToAddress(rsQuery(0)) & """>")
					Else
						.Write(FormatIPAddress(ConvertIPNumberToAddress(rsQuery(0)), strRegistry, strClass) & vbcrlf)
						.Write("<input type=hidden name=data1 value=""" & ConvertIPNumberToAddress(rsQuery(0)) & """>")
					End If
					.Write("</td>")
					.Write("<td align=right>" & rsQuery(2) & "</td>") & vbcrlf
					.Write("<td align=right>" & sngPercent & "</td>")
					.Write("<td>")
					If blnShowGraph = True Then
						Call GenerateGraph(rsQuery(2), intMaxNumber, strClass)
					End If
					.Write("</td>")
					.Write("<input type=hidden name=data2 value=""" & rsQuery(2) & """>")
					.Write("<input type=hidden name=data3 value=""" & sngPercent & """>")
					.Write("</tr>") & vbcrlf
				End With
				rsQuery.Movenext
			Loop
			
			If rsQuery.RecordCount > 0  Then
				With Response
					.Write("<tr class=total>")
					.Write("<td colspan=2 align=right>Total: </td>")
					.Write("<td align=right>" & intTotal & "</td>")
					.Write("<td align=right>&nbsp;</td>")
					.Write("<td align=right>&nbsp;</td>")
					.Write("</tr>")
					.Write("<input type=hidden name=total value=" & intTotal & ">")
					.Write("</table>") & vbcrlf
				End With
			End if
			
			rsQuery.Close : Set rsQuery = Nothing
			
			Call DisplayReportFooter()
			Call DisplayIPWhois()
			
		Case 7 : ' DIRECTORIES (PAGES & FILES)
			
			intTotal = CountPageViews("", datStart, datEnd)
			
			strSql = "SELECT " & SetItems(blnItems, intItems, True) & "pn_path, COUNT(pn_path) AS dc_pagecount FROM " &_
				"(SELECT pl_datetime, pn_path FROM " & strTablePrefix & "PageLog, " & strTablePrefix & "PageNames " &_
				"WHERE pl_pn_id = pn_id " &_
				"AND pl_datetime BETWEEN " & FormatDatabaseDate(datStart & " 00:00:00") &_
				" AND " & FormatDatabaseDate(datEnd & " 23:59:59") & ") dtPageLog " &_
				"GROUP BY pn_path "

			If strDatabaseType = "MYSQL" Then
				strSql = strSql & "ORDER BY dc_pagecount DESC, pn_path ASC " & SetItems(blnItems, intItems, False)
			Else
				strSql = strSql & "ORDER BY COUNT(pn_path) DESC, pn_path ASC"
			End If
				
			Set rsQuery = Server.CreateObject("ADODB.Recordset")
			
			If strDatabaseType = "MYSQL" Then
				rsQuery.CursorLocation = 3
			End If
			
			rsQuery.Open strSql, objConn, 3, 1, &H0000
			
			Call DisplayReportHeader(strReportGroup, strReportName, strDesc)
			
			If Not rsQuery.Eof Then
				Response.Write("<table border=0 cellpadding=4 cellspacing=0 width=""100%"">") & vbcrlf
				Response.Write("<tr><th></th><th align=left>Directory</th><th align=right>Count</th><th>%</th><th></th></tr>") & vbcrlf
				Response.Write("<input type=hidden name=data1 value=""Directory"">")
				Response.Write("<input type=hidden name=data2 value=""Count"">")
				Response.Write("<input type=hidden name=data3 value=""%"">")
				Response.Write("<input type=hidden name=cols value=3>")
			Else
				Response.Write("<p class=nodata>There is no data available for the report period selected.</p>") & vbcrlf
			End if
			Do While Not rsQuery.Eof

				If intTotal > 0 Then
					sngPercent = FormatPercent(rsQuery(1) / intTotal, 2)
				Else
					sngPercent = FormatPercent(0, 2)
				End If
			
				intRow = intRow + 1
				If (intRow Mod 2) = 1 Then
					strClass = "data"
				Else
					strClass = "dataalt"
				End If
			
				If intRow = 1 Then
					intMaxNumber = rsQuery(1)
				End If
			
				With Response
					.Write("<tr class=" & strClass & ">") & vbcrlf
					.Write("<td align=right>" & intRow & ".</td>")
					.Write("<td align=left>" & rsQuery(0) & "</td>")
					.Write("<td align=right>" & rsQuery(1) & "</td>")
					.Write("<td align=right>" & sngPercent & "</td>")
					.Write("<td align=left>")
					If blnShowGraph = True Then
						Call GenerateGraph(rsQuery(1), intMaxNumber, strClass)
					End If
					.Write("</td>")
					.Write("</tr>") & vbcrlf
					.Write("<input type=hidden name=data1 value=""" & FormatExportData(rsQuery(0)) & """>")
					.Write("<input type=hidden name=data2 value=" & rsQuery(1) & ">")
					.Write("<input type=hidden name=data3 value=""" & sngPercent & """>") & vbcrlf
				End With
				rsQuery.Movenext
			Loop
			
			If rsQuery.RecordCount > 0  Then
				With Response
					.Write("<tr class=total>")
					.Write("<td colspan=2 align=right>Total: </td>")
					.Write("<td align=right>" & intTotal & "</td>")
					.Write("<td align=right>&nbsp;</td>")
					.Write("<td align=right>&nbsp;</td>")
					.Write("</tr>")
					.Write("<input type=hidden name=total value=" & intTotal & ">")
					.Write("</table>") & vbcrlf
				End With
			End if
			
			rsQuery.Close : Set rsQuery = Nothing
			
			Call DisplayReportFooter()
		
		Case 8 : ' FILE TYPES (PAGES & FILES)
			
			intTotal = CountFileTypes("", datStart, datEnd)
			
			strSql = "SELECT " & SetItems(blnItems, intItems, True) & "pn_extension, COUNT(pn_extension) AS dc_pagecount FROM " &_
				"(SELECT pl_datetime, pn_extension FROM " & strTablePrefix & "PageLog, " & strTablePrefix & "PageNames " &_
				"WHERE pl_pn_id = pn_id " &_
				"AND pl_datetime BETWEEN " & FormatDatabaseDate(datStart & " 00:00:00") &_
				" AND " & FormatDatabaseDate(datEnd & " 23:59:59") & " " &_
				"AND pn_extension <> '') dtPageLog " &_
				"GROUP BY pn_extension "

			If strDatabaseType = "MYSQL" Then
				strSql = strSql & "ORDER BY dc_pagecount DESC, pn_extension ASC " & SetItems(blnItems, intItems, False)
			Else
				strSql = strSql & "ORDER BY COUNT(pn_extension) DESC, pn_extension ASC"
			End If

			Set rsQuery = Server.CreateObject("ADODB.Recordset")

			If strDatabaseType = "MYSQL" Then
				rsQuery.CursorLocation = 3
			End If
			
			rsQuery.Open strSql, objConn, 3, 1, &H0000
			
			Call DisplayReportHeader(strReportGroup, strReportName, strDesc)
			
			If Not rsQuery.Eof Then
				Response.Write("<table border=0 cellpadding=4 cellspacing=0 width=""100%"">") & vbcrlf
				Response.Write("<tr><th></th><th align=left>File Type</th><th align=right>Count</th><th>%</th><th></th></tr>") & vbcrlf
				Response.Write("<input type=hidden name=data1 value=""File Types"">")
				Response.Write("<input type=hidden name=data2 value=""Count"">")
				Response.Write("<input type=hidden name=data3 value=""%"">")
				Response.Write("<input type=hidden name=cols value=3>")
			Else
				Response.Write("<p class=nodata>There is no data available for the report period selected.</p>") & vbcrlf
			End if
			Do While Not rsQuery.Eof

				If intTotal > 0 Then
					sngPercent = FormatPercent(rsQuery(1) / intTotal, 2)
				Else
					sngPercent = FormatPercent(0, 2)
				End If
			
				intRow = intRow + 1
				If (intRow Mod 2) = 1 Then
					strClass = "data"
				Else
					strClass = "dataalt"
				End If
			
				If intRow = 1 Then
					intMaxNumber = rsQuery(1)
				End If
			
				With Response
					.Write("<tr class=" & strClass & ">") & vbcrlf
					.Write("<td align=right>" & intRow & ".</td>")
					.Write("<td align=left>" & rsQuery(0) & "</td>")
					.Write("<td align=right>" & rsQuery(1) & "</td>")
					.Write("<td align=right>" & sngPercent & "</td>")
					.Write("<td align=left>")
					If blnShowGraph = True Then
						Call GenerateGraph(rsQuery(1), intMaxNumber, strClass)
					End If
					.Write("</td>")
					.Write("</tr>") & vbcrlf
					.Write("<input type=hidden name=data1 value=""" & FormatExportData(rsQuery(0)) & """>")
					.Write("<input type=hidden name=data2 value=" & rsQuery(1) & ">")
					.Write("<input type=hidden name=data3 value=""" & sngPercent & """>") & vbcrlf
				End With
				rsQuery.Movenext
			Loop
			
			If rsQuery.RecordCount > 0  Then
				With Response
					.Write("<tr class=total>")
					.Write("<td colspan=2 align=right>Total: </td>")
					.Write("<td align=right>" & intTotal & "</td>")
					.Write("<td align=right>&nbsp;</td>")
					.Write("<td align=right>&nbsp;</td>")
					.Write("</tr>")
					.Write("<input type=hidden name=total value=" & intTotal & ">")
					.Write("</table>") & vbcrlf
				End With
			End if
			
			rsQuery.Close : Set rsQuery = Nothing
			
			Call DisplayReportFooter()
		
		Case 9 : ' DAILY UNIQUE VISITORS (VISITORS)
			
			intTotal 		= CountDailyVisitors(datStart, datEnd)
			If blnShowGraph = True Then
				intMaxNumber 	= GetMaxVisitors()
			End If
			datCurrent		= datStart
			
			strSql	= "SELECT LogYear, LogMonth, LogDay, COUNT(s_ip) " &_
				"FROM (SELECT DISTINCT YEAR(pl_datetime) AS LogYear, " &_
				"MONTH(pl_datetime) AS LogMonth,  " &_
				"DAY(pl_datetime) AS LogDay, s_ip FROM (" &_
				"SELECT pl_datetime, s_ip FROM " & strTablePrefix & "PageLog, " & strTablePrefix & "Sessions " &_
				"WHERE pl_s_id = s_id) dtPageLog " &_
				"WHERE pl_datetime BETWEEN " & FormatDatabaseDate(datStart & " 00:00:00") &_
				" AND " & FormatDatabaseDate(datEnd & " 23:59:59") &_
				") dtSessions GROUP BY LogYear, LogMonth, LogDay " &_
				"ORDER BY LogYear ASC, LogMonth ASC, LogDay ASC"

			Set rsQuery = Server.CreateObject("ADODB.Recordset")
			
			If strDatabaseType = "MYSQL" Then
				rsQuery.CursorLocation = 3
			End If
			
			rsQuery.Open strSql, objConn, 3, 1, &H0000
			
			Call DisplayReportHeader(strReportGroup, strReportName, strDesc)
			
			With Response
				.Write("<table border=0 cellpadding=4 cellspacing=0 width=""100%"">") & vbcrlf
				.Write("<tr><th>Day</th><th>Count</th><th>%</th><th></th></tr>") & vbcrlf
				.Write("<input type=hidden name=data1 value=""Day"">")
				.Write("<input type=hidden name=data2 value=""Count"">")
				.Write("<input type=hidden name=data3 value=""%"">")
				.Write("<input type=hidden name=cols value=3>")
			End With
			
			For intDayLoop = 1 to DateDiff("d", datStart, datEnd) + 1
			
				If Not rsQuery.Eof Then

					If DateDiff("d", datCurrent, DateSerial(rsQuery(0), rsQuery(1), rsQuery(2))) > 0 Then
						intCount = 0
						blnMovenext = False
					Else
						intCount = rsQuery(3)
						blnMovenext = True
					End If
				Else
					intCount = 0
					blnMovenext = False
				End If

				If intTotal > 0 Then
					sngPercent = FormatPercent(intCount / intTotal, 2)
				Else
					sngPercent = FormatPercent(0, 2)
				End If
			
				intRow = intRow + 1
				If (intRow Mod 2) = 1 Then
					strClass = "data"
				Else
					strClass = "dataalt"
				End If
			
				With Response
					.Write("<tr class=" & strClass & ">") & vbcrlf
					.Write("<td align=left>" & FormatDisplayDate(datCurrent, strShortDate) & "</td>")
					.Write("<td align=right>" & intCount & "</td>")
					.Write("<td align=right>" & sngPercent & "</td>")
					.Write("<td align=left>")
					If blnShowGraph = True Then
						Call GenerateGraph(intCount, intMaxNumber, strClass)
					End If
					.Write("</td>")
					.Write("</tr>") & vbcrlf
					.Write("<input type=hidden name=data1 value=""" & FormatDisplayDate(datCurrent, strShortDate) & """>")
					.Write("<input type=hidden name=data2 value=" & intCount & ">")
					.Write("<input type=hidden name=data3 value=""" & sngPercent & """>")
				End With
				
				datCurrent = DateAdd("d", 1, datCurrent)
				
				If blnMovenext = True Then
					rsQuery.Movenext
				End If
			Next
			
			With Response
					.Write("<tr class=total>")
					.Write("<td align=right>Total: </td>")
					.Write("<td align=right>" & intTotal & "</td>")
					.Write("<td align=right>&nbsp;</td>")
					.Write("<td align=right>&nbsp;</td>")
					.Write("</tr>")
					.Write("<input type=hidden name=total value=" & intTotal & ">")
					.Write("</table><br>") & vbcrlf
			End With
			
			rsQuery.Close : Set rsQuery = Nothing
			
			Call DisplayReportFooter()
		
		Case 10 : ' DAILY VISITS (VISITORS)
			
			intTotal 		= CountDailyVisits(datStart, datEnd)
			If blnShowGraph = True Then
				intMaxNumber 	= GetMaxVisits()
			End If
			datCurrent		= datStart
			
			strSql	= "SELECT LogYear, LogMonth, LogDay, COUNT(pl_s_id) " &_
				"FROM (SELECT DISTINCT YEAR(pl_datetime) AS LogYear, " &_
				"MONTH(pl_datetime) AS LogMonth,  DAY(pl_datetime) AS LogDay, pl_s_id " &_
				"FROM " & strTablePrefix & "PageLog " &_
				"WHERE pl_datetime BETWEEN " & FormatDatabaseDate(datStart & " 00:00:00") &_
				" AND " & FormatDatabaseDate(datEnd & " 23:59:59") &_
				") dtPageLog GROUP BY LogYear, LogMonth, LogDay " &_
				"ORDER BY LogYear ASC, LogMonth ASC, LogDay ASC"

			Set rsQuery = Server.CreateObject("ADODB.Recordset")
			rsQuery.Open strSql, objConn, 3, 1, &H0000
			
			Call DisplayReportHeader(strReportGroup, strReportName, strDesc)
			
			With Response
				.Write("<table border=0 cellpadding=4 cellspacing=0 width=""100%"">") & vbcrlf
				.Write("<tr><th>Day</th><th>Count</th><th>%</th><th></th></tr>") & vbcrlf
				.Write("<input type=hidden name=data1 value=""Day"">")
				.Write("<input type=hidden name=data2 value=""Count"">")
				.Write("<input type=hidden name=data3 value=""%"">")
				.Write("<input type=hidden name=cols value=3>")
			End With
			
			For intDayLoop = 1 to DateDiff("d", datStart, datEnd) + 1
			
				If Not rsQuery.Eof Then

					If DateDiff("d", datCurrent, DateSerial(rsQuery(0), rsQuery(1), rsQuery(2))) > 0 Then
						intCount = 0
						blnMovenext = False
					Else
						intCount = rsQuery(3)
						blnMovenext = True
					End If
				Else
					intCount = 0
					blnMovenext = False
				End If

				If intTotal > 0 Then
					sngPercent = FormatPercent(intCount / intTotal, 2)
				Else
					sngPercent = FormatPercent(0, 2)
				End If
			
				intRow = intRow + 1
				If (intRow Mod 2) = 1 Then
					strClass = "data"
				Else
					strClass = "dataalt"
				End If
			
				With Response
					.Write("<tr class=" & strClass & ">") & vbcrlf
					.Write("<td align=left>" & FormatDisplayDate(datCurrent, strShortDate) & "</td>")
					.Write("<td align=right>" & intCount & "</td>")
					.Write("<td align=right>" & sngPercent & "</td>")
					.Write("<td align=left>")
					If blnShowGraph = True Then
						Call GenerateGraph(intCount, intMaxNumber, strClass)
					End If
					.Write("</td>")
					.Write("</tr>") & vbcrlf
					.Write("<input type=hidden name=data1 value=""" & FormatDisplayDate(datCurrent, strShortDate) & """>")
					.Write("<input type=hidden name=data2 value=" & intCount & ">")
					.Write("<input type=hidden name=data3 value=""" & sngPercent & """>")
				End With
				
				datCurrent = DateAdd("d", 1, datCurrent)
				
				If blnMovenext = True Then
					rsQuery.Movenext
				End If
			Next
			
			With Response
					.Write("<tr class=total>")
					.Write("<td align=right>Total: </td>")
					.Write("<td align=right>" & intTotal & "</td>")
					.Write("<td align=right>&nbsp;</td>")
					.Write("<td align=right>&nbsp;</td>")
					.Write("</tr>")
					.Write("<input type=hidden name=total value=" & intTotal & ">")
					.Write("</table><br>") & vbcrlf
			End With
			
			rsQuery.Close : Set rsQuery = Nothing
			
			Call DisplayReportFooter()

		Case 11 : ' VISITS BY IP ADDRESS (VISITORS)
			
			intTotal = CountVisits(datStart, datEnd)

			strSql = "SELECT " & SetItems(blnItems, intItems, True) & " s_ip, n_value AS s_hostname, " &_
				"COUNT(s_id) AS dc_count, s_country " &_
				"FROM (SELECT DISTINCT s_id, s_ip, s_hostname, s_country " &_
				"FROM " & strTablePrefix & "PageLog, " & strTablePrefix & "Sessions " &_
				"WHERE pl_s_id = s_id " &_
				"AND pl_datetime BETWEEN " & FormatDatabaseDate(datStart & " 00:00:00") &_
				" AND " & FormatDatabaseDate(datEnd & " 23:59:59") & ") dt_PageLog"
				
			If strDatabaseType = "MYSQL" Then
				strSql = strSql & " LEFT JOIN " & strTablePrefix & "Names ON s_hostname = n_id " &_
					"GROUP BY s_ip, s_hostname, s_country " &_
					"ORDER BY dc_count DESC, s_ip ASC " & SetItems(blnItems, intItems, False)
			Else
				strSql = strSql & " LEFT JOIN " & strTablePrefix & "Names " &_
					"ON dt_PageLog.s_hostname = " & strTablePrefix & "Names.n_id " &_
					"GROUP BY s_ip, n_value, s_country " &_
					"ORDER BY COUNT(s_id) DESC, s_ip ASC"
			End If
				
			Set rsQuery = Server.CreateObject("ADODB.Recordset")
			
			If strDatabaseType = "MYSQL" Then
				rsQuery.CursorLocation = 3
			End If
			
			rsQuery.Open strSql, objConn, 3, 1, &H0000
			
			Call DisplayReportHeader(strReportGroup, strReportName, strDesc)
			
			If Not rsQuery.Eof Then
				Response.Write("<table border=0 cellpadding=4 cellspacing=0 width=""100%"">") & vbcrlf
				Response.Write("<tr><th></th><th>IP Address</th><th>Visits</th><th>%</th><th></th></tr>") & vbcrlf
				Response.Write("<input type=hidden name=data1 value=""IP Address"">")
				Response.Write("<input type=hidden name=data2 value=""Visits"">")
				Response.Write("<input type=hidden name=data3 value=""%"">")
				Response.Write("<input type=hidden name=cols value=3>")
			Else
				Response.Write("<p class=nodata>There is no data available for the report period selected.</p>") & vbcrlf
			End if
			
			Do While Not rsQuery.Eof

				If intTotal > 0 Then
					sngPercent = FormatPercent(rsQuery(2) / intTotal, 2)
				Else
					sngPercent = FormatPercent(0, 2)
				End If
			
				intRow = intRow + 1
				If (intRow Mod 2) = 1 Then
					strClass = "data"
				Else
					strClass = "dataalt"
				End If
			
				If intRow = 1 Then
					intMaxNumber = rsQuery(2)
				End If
			
				strRegistry = ChooseIPWhois(rsQuery(3))
			
				With Response
					.Write("<tr class=" & strClass & ">") & vbcrlf
					.Write("<td align=right>" & intRow & ".</td>")
					.Write("<td align=left>") & vbcrlf
					If rsQuery(1) <> "" Then
						.Write(rsQuery(1) & "<br>")
						.Write(FormatIPAddress(ConvertIPNumberToAddress(rsQuery(0)), strRegistry, strClass) & vbcrlf)
						.Write("<input type=hidden name=data1 value=""" & rsQuery(1) & " - " & ConvertIPNumberToAddress(rsQuery(0)) & """>")
					Else
						.Write(FormatIPAddress(ConvertIPNumberToAddress(rsQuery(0)), strRegistry, strClass) & vbcrlf)
						.Write("<input type=hidden name=data1 value=""" & ConvertIPNumberToAddress(rsQuery(0)) & """>")
					End If
					.Write("</td>")
					.Write("<td align=right>" & rsQuery(2) & "</td>") & vbcrlf
					.Write("<td align=right>" & sngPercent & "</td>")
					If blnShowGraph = True Then
						.Write("<td align=left>")
						Call GenerateGraph(rsQuery(2), intMaxNumber, strClass)
						.Write("</td>")
					End If
					.Write("<input type=hidden name=data2 value=""" & rsQuery(2) & """>")
					.Write("<input type=hidden name=data3 value=""" & sngPercent & """>")
					.Write("</tr>") & vbcrlf
				End With
				rsQuery.Movenext
			Loop
			
			If rsQuery.RecordCount > 0  Then
				With Response
					.Write("<tr class=total>")
					.Write("<td colspan=2 align=right>Total: </td>")
					.Write("<td align=right>" & intTotal & "</td>")
					.Write("<td align=right>&nbsp;</td>")
					.Write("<td align=right>&nbsp;</td>")
					.Write("</tr>")
					.Write("<input type=hidden name=total value=" & intTotal & ">")
					.Write("</table>") & vbcrlf
				End With
			End if
			
			rsQuery.Close : Set rsQuery = Nothing
			
			Call DisplayReportFooter()
			Call DisplayIPWhois()
		
		Case 12 : ' ENTRY PAGES (PAGES & FILES)
			
			intTotal = CountVisits(datStart, datEnd)

			strSql = "SELECT " & SetItems(blnItems, intItems, True) & "pn_url, COUNT(pn_url) AS dc_pagecount, pn_label " &_
				"FROM " & strTablePrefix & "PageNames, " & strTablePrefix & "PageLog,  " &_
				"(SELECT MIN(pl_datetime) AS dl_datetime, pl_s_id AS dl_s_id " &_
				"FROM " & strTablePrefix & "PageLog " &_
				"WHERE pl_datetime BETWEEN " & FormatDatabaseDate(datStart & " 00:00:00") &_
				" AND " & FormatDatabaseDate(datEnd & " 23:59:59") & " " &_
				"GROUP BY pl_s_id) dtPageLog " &_
				"WHERE pn_id = pl_pn_id " &_
				"AND pl_s_id = dl_s_id " &_
				"AND dl_datetime = pl_datetime " &_
				"GROUP BY pn_url, pn_label "
				
			If strDatabaseType = "MYSQL" Then
				strSql = strSql & "ORDER BY dc_pagecount DESC, pn_page ASC " & SetItems(blnItems, intItems, False)
			Else
				strSql = strSql & "ORDER BY COUNT(pn_url) DESC, pn_url ASC"
			End If
				
			Set rsQuery = Server.CreateObject("ADODB.Recordset")
			
			If strDatabaseType = "MYSQL" Then
				rsQuery.CursorLocation = 3
			End If
			
			rsQuery.Open strSql, objConn, 3, 1, &H0000
			
			Call DisplayReportHeader(strReportGroup, strReportName, strDesc)
			
			If Not rsQuery.Eof Then
				Response.Write("<table border=0 cellpadding=4 cellspacing=0 width=""100%"">") & vbcrlf
				Response.Write("<tr><th></th><th align=left>Url</th><th align=right>Count</th><th>%</th><th></th></tr>") & vbcrlf
				Response.Write("<input type=hidden name=data1 value=""Url"">")
				Response.Write("<input type=hidden name=data2 value=""Count"">")
				Response.Write("<input type=hidden name=data3 value=""%"">")
				Response.Write("<input type=hidden name=cols value=3>")
			Else
				Response.Write("<p class=nodata>There is no data available for the report period selected.</p>") & vbcrlf
			End if
			
			Do While Not rsQuery.Eof

				If intTotal > 0 Then
					sngPercent = FormatPercent(rsQuery(1) / intTotal, 2)
				Else
					sngPercent = FormatPercent(0, 2)
				End If
			
				intRow = intRow + 1
				If (intRow Mod 2) = 1 Then
					strClass = "data"
				Else
					strClass = "dataalt"
				End If
			
				If intRow = 1 Then
					intMaxNumber = rsQuery(1)
				End If
				
				With Response
					.Write("<tr class=" & strClass & ">") & vbcrlf
					.Write("<td align=right>" & intRow & ".</td>")
					.Write("<td align=left>")
					If rsQuery(2) <> "" Then
						.Write(rsQuery(2) & "<br>")
					End If
					.Write(FormatLink(rsQuery(0), strClass) & "</td>")
					.Write("<td align=right>" & rsQuery(1) & "</td>")
					.Write("<td align=right>" & sngPercent & "</td>")
					.Write("<td align=left>")
					If blnShowGraph = True Then
						Call GenerateGraph(rsQuery(1), intMaxNumber, strClass)
					End If
					.Write("</td>")
					.Write("</tr>") & vbcrlf
					If rsQuery(2) <> "" Then
						.Write("<input type=hidden name=data1 value=""" & FormatExportData(rsQuery(2) & " (" & rsQuery(0)) & ")"">")
					Else
						.Write("<input type=hidden name=data1 value=""" & FormatExportData(rsQuery(0)) & """>")
					End If
					.Write("<input type=hidden name=data2 value=" & rsQuery(1) & ">")
					.Write("<input type=hidden name=data3 value=""" & sngPercent & """>") & vbcrlf
				End With

				rsQuery.Movenext
			Loop
			
			If rsQuery.RecordCount > 0  Then
				With Response
					.Write("<tr class=total>")
					.Write("<td colspan=2 align=right>Total: </td>")
					.Write("<td align=right>" & intTotal & "</td>")
					.Write("<td align=right>&nbsp;</td>")
					.Write("<td align=right>&nbsp;</td>")
					.Write("</tr>")
					.Write("<input type=hidden name=total value=" & intTotal & ">")
					.Write("</table>") & vbcrlf
				End With
			End if
			
			rsQuery.Close : Set rsQuery = Nothing
			
			Call DisplayReportFooter()
		
		Case 13 : ' EXIT PAGES (PAGES & FILES)
			
			intTotal = CountVisits(datStart, datEnd)

			strSql = "SELECT " & SetItems(blnItems, intItems, True) & " pn_url, COUNT(pn_url) AS dc_pagecount, pn_label " &_
				"FROM " & strTablePrefix & "PageNames, " & strTablePrefix & "PageLog,  " &_
				"(SELECT MAX(pl_datetime) AS dl_datetime, pl_s_id AS dl_s_id " &_
				"FROM " & strTablePrefix & "PageLog " &_
				"WHERE pl_datetime BETWEEN " & FormatDatabaseDate(datStart & " 00:00:00") &_
				" AND " & FormatDatabaseDate(datEnd & " 23:59:59") & " " &_
				"GROUP BY pl_s_id) dtPageLog " &_
				"WHERE pn_id = pl_pn_id " &_
				"AND pl_s_id = dl_s_id " &_
				"AND dl_datetime = pl_datetime " &_
				"GROUP BY pn_url, pn_label "
				
			If strDatabaseType = "MYSQL" Then
				strSql = strSql & "ORDER BY dc_pagecount DESC, pn_page ASC " & SetItems(blnItems, intItems, False)
			Else
				strSql = strSql & "ORDER BY COUNT(pn_url) DESC, pn_url ASC"
			End If

			Set rsQuery = Server.CreateObject("ADODB.Recordset")
			
			If strDatabaseType = "MYSQL" Then
				rsQuery.CursorLocation = 3
			End If
			
			rsQuery.Open strSql, objConn, 3, 1, &H0000
			
			Call DisplayReportHeader(strReportGroup, strReportName, strDesc)
			
			If Not rsQuery.Eof Then
				Response.Write("<table border=0 cellpadding=4 cellspacing=0 width=""100%"">") & vbcrlf
				Response.Write("<tr><th></th><th align=left>Url</th><th align=right>Count</th><th>%</th><th></th></tr>") & vbcrlf
				Response.Write("<input type=hidden name=data1 value=""Url"">")
				Response.Write("<input type=hidden name=data2 value=""Count"">")
				Response.Write("<input type=hidden name=data3 value=""%"">")
				Response.Write("<input type=hidden name=cols value=3>")
			Else
				Response.Write("<p class=nodata>There is no data available for the report period selected.</p>") & vbcrlf
			End if
			
			Do While Not rsQuery.Eof

				If intTotal > 0 Then
					sngPercent = FormatPercent(rsQuery(1) / intTotal, 2)
				Else
					sngPercent = FormatPercent(0, 2)
				End If
			
				intRow = intRow + 1
				If (intRow Mod 2) = 1 Then
					strClass = "data"
				Else
					strClass = "dataalt"
				End If
			
				If intRow = 1 Then
					intMaxNumber = rsQuery(1)
				End If
				
				With Response
					.Write("<tr class=" & strClass & ">") & vbcrlf
					.Write("<td align=right>" & intRow & ".</td>")
					.Write("<td align=left>")
					If rsQuery(2) <> "" Then
						.Write(rsQuery(2) & "<br>")
					End If
					.Write(FormatLink(rsQuery(0), strClass) & "</td>")
					.Write("<td align=right>" & rsQuery(1) & "</td>")
					.Write("<td align=right>" & sngPercent & "</td>")
					.Write("<td align=left>")
					If blnShowGraph = True Then
						Call GenerateGraph(rsQuery(1), intMaxNumber, strClass)
					End If
					.Write("</td>")
					.Write("</tr>") & vbcrlf
					If rsQuery(2) <> "" Then
						.Write("<input type=hidden name=data1 value=""" & FormatExportData(rsQuery(2) & " (" & rsQuery(0)) & ")"">")
					Else
						.Write("<input type=hidden name=data1 value=""" & FormatExportData(rsQuery(0)) & """>")
					End If
					.Write("<input type=hidden name=data2 value=" & rsQuery(1) & ">")
					.Write("<input type=hidden name=data3 value=""" & sngPercent & """>") & vbcrlf
				End With

				rsQuery.Movenext
			Loop
			
			If rsQuery.RecordCount > 0  Then
				With Response
					.Write("<tr class=total>")
					.Write("<td colspan=2 align=right>Total: </td>")
					.Write("<td align=right>" & intTotal & "</td>")
					.Write("<td align=right>&nbsp;</td>")
					.Write("<td align=right>&nbsp;</td>")
					.Write("</tr>")
					.Write("<input type=hidden name=total value=" & intTotal & ">")
					.Write("</table>") & vbcrlf
				End With
			End if
			
			rsQuery.Close : Set rsQuery = Nothing
			
			Call DisplayReportFooter()
		
		Case 14 : ' DOMAINS (REFERRERS)
			
			intTotal = CountReferrers("Domain", datStart, datEnd)

			strSql = "SELECT " & SetItems(blnItems, intItems, True) & " rn_domain, rn_extension, COUNT(rn_domain) AS dc_count " &_
				"FROM (SELECT rn_domain, rn_extension " &_
				"FROM " & strTablePrefix & "Referrers, " & strTablePrefix & "ReferrerNames, " & strTablePrefix & "PageLog " &_
				"WHERE r_rn_id = rn_id AND pl_r_id = r_id "
			
			If blnSiteAliases = True Then
				strSql = strSql & "AND rn_host NOT IN (" & FormatSiteAliases(strSiteAliases) & ") "
			End If
			
			strSql = strSql & "AND pl_datetime BETWEEN " & FormatDatabaseDate(datStart & " 00:00:00") &_
				" AND " & FormatDatabaseDate(datEnd & " 23:59:59") & " " &_
				"AND rn_domain <> '') dt_referrers "
			
			strSql = strSql & "GROUP BY rn_domain, rn_extension "
			
			If strDatabaseType = "MYSQL" Then
				strSql = strSql & "ORDER BY dc_count DESC, rn_domain ASC " & SetItems(blnItems, intItems, False)
			Else
				strSql = strSql & "ORDER BY COUNT(rn_domain) DESC, rn_domain ASC"
			End If
				
			Set rsQuery = Server.CreateObject("ADODB.Recordset")
			
			If strDatabaseType = "MYSQL" Then
				rsQuery.CursorLocation = 3
			End If
			
			rsQuery.Open strSql, objConn, 3, 1, &H0000
			
			Call DisplayReportHeader(strReportGroup, strReportName, strDesc)
			
			If Not rsQuery.Eof Then
				Response.Write("<table border=0 cellpadding=4 cellspacing=0 width=""100%"">") & vbcrlf
				Response.Write("<tr><th></th><th>Domain</th><th>Count</th><th>%</th><th></th></tr>") & vbcrlf
				Response.Write("<input type=hidden name=data1 value=""Domain"">")
				Response.Write("<input type=hidden name=data2 value=""Count"">")
				Response.Write("<input type=hidden name=data3 value=""%"">")
				Response.Write("<input type=hidden name=cols value=3>")
			Else
				Response.Write("<p class=nodata>There is no data available for the report period selected.</p>") & vbcrlf
			End if
			Do While Not rsQuery.Eof

				If intTotal > 0 Then
					sngPercent = FormatPercent(rsQuery(2) / intTotal, 2)
				Else
					sngPercent = FormatPercent(0, 2)
				End If
			
				intRow = intRow + 1
				If (intRow Mod 2) = 1 Then
					strClass = "data"
				Else
					strClass = "dataalt"
				End If
			
				If intRow = 1 Then
					intMaxNumber = rsQuery(2)
				End If
			
				With Response
					.Write("<tr class=" & strClass & ">") & vbcrlf
					.Write("<td align=right>" & intRow & ".</td>")
					.Write("<td align=left>" & DisplayReferrerLink(rsQuery(0), rsQuery(0), strClass) & "</td>")
					.Write("<td align=right>" & rsQuery(2) & "</td>")
					.Write("<td align=right>" & sngPercent & "</td>")
					.Write("<td align=left>")
					If blnShowGraph = True Then
						Call GenerateGraph(rsQuery(2), intMaxNumber, strClass)
					End If
					.Write("</td>")
					.Write("</tr>") & vbcrlf
					.Write("<input type=hidden name=data1 value=""" & rsQuery(0) & """>")
					.Write("<input type=hidden name=data2 value=" & rsQuery(2) & ">")
					.Write("<input type=hidden name=data3 value=""" & sngPercent & """>") & vbcrlf
				End With
				rsQuery.Movenext
			Loop
			
			If rsQuery.RecordCount > 0  Then
				With Response
					.Write("<tr class=total>")
					.Write("<td colspan=2 align=right>Total: </td>")
					.Write("<td align=right>" & intTotal & "</td>")
					.Write("<td align=right>&nbsp;</td>")
					.Write("<td align=right>&nbsp;</td>")
					.Write("</tr>")
					.Write("<input type=hidden name=total value=" & intTotal & ">")
					.Write("</table>") & vbcrlf
				End With
			End if
			
			rsQuery.Close : Set rsQuery = Nothing
			
			Call DisplayReportFooter()
			
		Case 15 : ' HOSTS (REFERRERS)
			
			intTotal = CountReferrers("Host", datStart, datEnd)
			
			strSql = "SELECT " & SetItems(blnItems, intItems, True) & " rn_host, rn_domain, rn_extension, COUNT(rn_host) AS dc_count " &_
				"FROM (SELECT rn_host, rn_domain, rn_extension " &_
				"FROM " & strTablePrefix & "Referrers, " & strTablePrefix & "ReferrerNames, " & strTablePrefix & "PageLog " &_
				"WHERE r_rn_id = rn_id AND pl_r_id = r_id "
			
			If blnSiteAliases = True Then
				strSql = strSql & "AND rn_host NOT IN (" & FormatSiteAliases(strSiteAliases) & ") "
			End If
			
			strSql = strSql & "AND pl_datetime BETWEEN " & FormatDatabaseDate(datStart & " 00:00:00") &_
				" AND " & FormatDatabaseDate(datEnd & " 23:59:59") & " " &_
				"AND rn_host <> '') dt_referrers "
			
			strSql = strSql & "GROUP BY rn_host, rn_domain, rn_extension "
			
			If strDatabaseType = "MYSQL" Then
				strSql = strSql & "ORDER BY dc_count DESC, rn_host ASC " & SetItems(blnItems, intItems, False)
			Else
				strSql = strSql & "ORDER BY COUNT(rn_host) DESC, rn_host ASC"
			End If
				
			Set rsQuery = Server.CreateObject("ADODB.Recordset")
			
			If strDatabaseType = "MYSQL" Then
				rsQuery.CursorLocation = 3
			End If
			
			rsQuery.Open strSql, objConn, 3, 1, &H0000
			
			Call DisplayReportHeader(strReportGroup, strReportName, strDesc)
			
			If Not rsQuery.Eof Then
				Response.Write("<table border=0 cellpadding=4 cellspacing=0 width=""100%"">") & vbcrlf
				Response.Write("<tr><th></th><th>Host</th><th>Count</th><th>%</th><th></th></tr>") & vbcrlf
				Response.Write("<input type=hidden name=data1 value=""Host"">")
				Response.Write("<input type=hidden name=data2 value=""Count"">")
				Response.Write("<input type=hidden name=data3 value=""%"">")
				Response.Write("<input type=hidden name=cols value=3>")
			Else
				Response.Write("<p class=nodata>There is no data available for the report period selected.</p>") & vbcrlf
			End if
			Do While Not rsQuery.Eof

				If intTotal > 0 Then
					sngPercent = FormatPercent(rsQuery(3) / intTotal, 2)
				Else
					sngPercent = FormatPercent(0, 2)
				End If
			
				intRow = intRow + 1
				If (intRow Mod 2) = 1 Then
					strClass = "data"
				Else
					strClass = "dataalt"
				End If
			
				If intRow = 1 Then
					intMaxNumber = rsQuery(3)
				End If
			
				With Response
					.Write("<tr class=" & strClass & ">") & vbcrlf
					.Write("<td align=right>" & intRow & ".</td>")
					.Write("<td align=left>" & DisplayReferrerLink(rsQuery(0), rsQuery(0), strClass) & "</td>")
					.Write("<td align=right>" & rsQuery(3) & "</td>")
					.Write("<td align=right>" & sngPercent & "</td>")
					.Write("<td align=left>")
					If blnShowGraph = True Then
						Call GenerateGraph(rsQuery(3), intMaxNumber, strClass)
					End If
					.Write("</td>")
					.Write("</tr>") & vbcrlf
					.Write("<input type=hidden name=data1 value=""" & rsQuery(0) & """>")
					.Write("<input type=hidden name=data2 value=" & rsQuery(3) & ">")
					.Write("<input type=hidden name=data3 value=""" & sngPercent & """>") & vbcrlf
				End With
				rsQuery.Movenext
			Loop
			
			If rsQuery.RecordCount > 0  Then
				With Response
					.Write("<tr class=total>")
					.Write("<td colspan=2 align=right>Total: </td>")
					.Write("<td align=right>" & intTotal & "</td>")
					.Write("<td align=right>&nbsp;</td>")
					.Write("<td align=right>&nbsp;</td>")
					.Write("</tr>")
					.Write("<input type=hidden name=total value=" & intTotal & ">")
					.Write("</table>") & vbcrlf
				End With
			End if
			
			rsQuery.Close : Set rsQuery = Nothing
			
			Call DisplayReportFooter()
			
		Case 16 : ' PAGES (REFERRERS)
			
			intTotal = CountReferrers("Page", datStart, datEnd)

			strSql = "SELECT " & SetItems(blnItems, intItems, True) & " rn_page, COUNT(rn_page) AS dc_count " &_
				"FROM (SELECT rn_page " &_
				"FROM " & strTablePrefix & "Referrers, " & strTablePrefix & "ReferrerNames, " & strTablePrefix & "PageLog " &_
				"WHERE r_rn_id = rn_id AND pl_r_id = r_id "
			
			If blnSiteAliases = True Then
				strSql = strSql & "AND rn_host NOT IN (" & FormatSiteAliases(strSiteAliases) & ") "
			End If
			
			strSql = strSql & "AND pl_datetime BETWEEN " & FormatDatabaseDate(datStart & " 00:00:00") &_
				" AND " & FormatDatabaseDate(datEnd & " 23:59:59") & " " &_
				"AND rn_page <> '') dt_referrers "
			
			strSql = strSql & "GROUP BY rn_page "
			
			If strDatabaseType = "MYSQL" Then
				strSql = strSql & "ORDER BY dc_count DESC, rn_page ASC " & SetItems(blnItems, intItems, False)
			Else
				strSql = strSql & "ORDER BY COUNT(rn_page) DESC, rn_page ASC"
			End If
				
			Set rsQuery = Server.CreateObject("ADODB.Recordset")
			
			If strDatabaseType = "MYSQL" Then
				rsQuery.CursorLocation = 3
			End If
			
			rsQuery.Open strSql, objConn, 3, 1, &H0000
			
			Call DisplayReportHeader(strReportGroup, strReportName, strDesc)
			
			If Not rsQuery.Eof Then
				Response.Write("<table border=0 cellpadding=4 cellspacing=0 width=""100%"">") & vbcrlf
				Response.Write("<tr><th></th><th>Page</th><th>Count</th><th>%</th><th></th></tr>") & vbcrlf
				Response.Write("<input type=hidden name=data1 value=""Page"">")
				Response.Write("<input type=hidden name=data2 value=""Count"">")
				Response.Write("<input type=hidden name=data3 value=""%"">")
				Response.Write("<input type=hidden name=cols value=3>")
			Else
				Response.Write("<p class=nodata>There is no data available for the report period selected.</p>") & vbcrlf
			End if
			Do While Not rsQuery.Eof

				If intTotal > 0 Then
					sngPercent = FormatPercent(rsQuery(1) / intTotal, 2)
				Else
					sngPercent = FormatPercent(0, 2)
				End If
			
				intRow = intRow + 1
				If (intRow Mod 2) = 1 Then
					strClass = "data"
				Else
					strClass = "dataalt"
				End If
			
				If intRow = 1 Then
					intMaxNumber = rsQuery(1)
				End If
				
				strUrl = TruncateUrl(rsQuery(0))
				
				With Response
					.Write("<tr class=" & strClass & ">") & vbcrlf
					.Write("<td align=right>" & intRow & ".</td>")
					.Write("<td align=left>" & DisplayReferrerLink(strUrl, rsQuery(0), strClass))
					.Write("</td>")
					.Write("<td align=right>" & rsQuery(1) & "</td>")
					.Write("<td align=right>" & sngPercent & "</td>")
					.Write("<td align=left>")
					If blnShowGraph = True Then
						Call GenerateGraph(rsQuery(1), intMaxNumber, strClass)
					End If
					.Write("</td>")
					.Write("</tr>") & vbcrlf
					.Write("<input type=hidden name=data1 value=""" & FormatExportData(rsQuery(0)) & """>")
					.Write("<input type=hidden name=data2 value=" & rsQuery(1) & ">")
					.Write("<input type=hidden name=data3 value=""" & sngPercent & """>") & vbcrlf
				End With
				rsQuery.Movenext
			Loop
			
			If rsQuery.RecordCount > 0  Then
				With Response
					.Write("<tr class=total>")
					.Write("<td colspan=2 align=right>Total: </td>")
					.Write("<td align=right>" & intTotal & "</td>")
					.Write("<td align=right>&nbsp;</td>")
					.Write("<td align=right>&nbsp;</td>")
					.Write("</tr>")
					.Write("<input type=hidden name=total value=" & intTotal & ">")
					.Write("</table>") & vbcrlf
				End With
			End if
			
			rsQuery.Close : Set rsQuery = Nothing
			
			Call DisplayReportFooter()
			
		Case 17 : ' Url (REFERRERS)
			
			intTotal = CountReferrers("Url", datStart, datEnd)
			
			strSql = "SELECT " & SetItems(blnItems, intItems, True) & " r_url, COUNT(r_url) AS dc_count " &_
				"FROM (SELECT r_url " &_
				"FROM " & strTablePrefix & "Referrers, " & strTablePrefix & "ReferrerNames, " & strTablePrefix & "PageLog " &_
				"WHERE r_rn_id = rn_id AND pl_r_id = r_id "
			
			If blnSiteAliases = True Then
				strSql = strSql & "AND rn_host NOT IN (" & FormatSiteAliases(strSiteAliases) & ") "
			End If
			
			strSql = strSql & "AND pl_datetime BETWEEN " & FormatDatabaseDate(datStart & " 00:00:00") &_
				" AND " & FormatDatabaseDate(datEnd & " 23:59:59") & " " &_
				"AND r_url <> '') dt_referrers "
			
			strSql = strSql & "GROUP BY r_url "
			
			If strDatabaseType = "MYSQL" Then
				strSql = strSql & "ORDER BY dc_count DESC, r_url ASC " & SetItems(blnItems, intItems, False)
			Else
				strSql = strSql & "ORDER BY COUNT(r_url) DESC, r_url ASC"
			End If
			
			Set rsQuery = Server.CreateObject("ADODB.Recordset")
			
			If strDatabaseType = "MYSQL" Then
				rsQuery.CursorLocation = 3
			End If
			
			rsQuery.Open strSql, objConn, 3, 1, &H0000
			
			Call DisplayReportHeader(strReportGroup, strReportName, strDesc)
			
			If Not rsQuery.Eof Then
				Response.Write("<table border=0 cellpadding=4 cellspacing=0 width=""100%"">") & vbcrlf
				Response.Write("<tr><th></th><th>Url</th><th>Count</th><th>%</th><th></th></tr>") & vbcrlf
				Response.Write("<input type=hidden name=data1 value=""Url"">")
				Response.Write("<input type=hidden name=data2 value=""Count"">")
				Response.Write("<input type=hidden name=data3 value=""%"">")
				Response.Write("<input type=hidden name=cols value=3>")
				Response.Write("<input type=hidden name=records value=" & rsQuery.RecordCount & ">")
			Else
				Response.Write("<p class=nodata>There is no data available for the report period selected.</p>") & vbcrlf
			End if
			Do While Not rsQuery.Eof

				If intTotal > 0 Then
					sngPercent = FormatPercent(rsQuery(1) / intTotal, 2)
				Else
					sngPercent = FormatPercent(0, 2)
				End If
			
				intRow = intRow + 1
				If (intRow Mod 2) = 1 Then
					strClass = "data"
				Else
					strClass = "dataalt"
				End If
			
				If intRow = 1 Then
					intMaxNumber = rsQuery(1)
				End If
			
				strUrl = TruncateUrl(rsQuery(0))
			
				With Response
					.Write("<tr class=" & strClass & ">") & vbcrlf
					.Write("<td align=right>" & intRow & ".</td>")
					.Write("<td align=left>" & DisplayReferrerLink(strUrl, rsQuery(0), strClass))
					.Write("</td>")
					.Write("<td align=right>" & rsQuery(1) & "</td>")
					.Write("<td align=right>" & sngPercent & "</td>")
					.Write("<td align=left>")
					If blnShowGraph = True Then
						Call GenerateGraph(rsQuery(1), intMaxNumber, strClass)
					End If
					.Write("</td>")
					.Write("</tr>") & vbcrlf
					.Write("<input type=hidden name=data1 value=""" & FormatExportData(rsQuery(0)) & """>")
					.Write("<input type=hidden name=data2 value=" & rsQuery(1) & ">")
					.Write("<input type=hidden name=data3 value=""" & sngPercent & """>") & vbcrlf
				End With
				rsQuery.Movenext
			Loop
			
			If rsQuery.RecordCount > 0  Then
				With Response
					.Write("<tr class=total>")
					.Write("<td colspan=2 align=right>Total: </td>")
					.Write("<td align=right>" & intTotal & "</td>")
					.Write("<td align=right>&nbsp;</td>")
					.Write("<td align=right>&nbsp;</td>")
					.Write("</tr>")
					.Write("<input type=hidden name=total value=" & intTotal & ">")
					.Write("</table>") & vbcrlf
				End With
			End if
			
			rsQuery.Close : Set rsQuery = Nothing
			
			Call DisplayReportFooter()
			
		Case 18 : ' EXTENSIONS (REFERRERS)
			
			intTotal = CountReferrers("Extension", datStart, datEnd)
			
			strSql = "SELECT " & SetItems(blnItems, intItems, True) & " rn_extension, COUNT(rn_extension) AS dc_count " &_
				"FROM (SELECT rn_extension " &_
				"FROM " & strTablePrefix & "Referrers, " & strTablePrefix & "ReferrerNames, " & strTablePrefix & "PageLog " &_
				"WHERE r_rn_id = rn_id AND pl_r_id = r_id "
			
			If blnSiteAliases = True Then
				strSql = strSql & "AND rn_host NOT IN (" & FormatSiteAliases(strSiteAliases) & ") "
			End If
			
			strSql = strSql & "AND pl_datetime BETWEEN " & FormatDatabaseDate(datStart & " 00:00:00") &_
				" AND " & FormatDatabaseDate(datEnd & " 23:59:59") & " " &_
				"AND rn_extension <> '') dt_referrers "
			
			strSql = strSql & "GROUP BY rn_extension "
			
			If strDatabaseType = "MYSQL" Then
				strSql = strSql & "ORDER BY dc_count DESC, rn_extension ASC " & SetItems(blnItems, intItems, False)
			Else
				strSql = strSql & "ORDER BY COUNT(rn_extension) DESC, rn_extension ASC"
			End If
				
			Set rsQuery = Server.CreateObject("ADODB.Recordset")
			
			If strDatabaseType = "MYSQL" Then
				rsQuery.CursorLocation = 3
			End If
			
			rsQuery.Open strSql, objConn, 3, 1, &H0000
			
			Call DisplayReportHeader(strReportGroup, strReportName, strDesc)
			
			If Not rsQuery.Eof Then
				Response.Write("<table border=0 cellpadding=4 cellspacing=0 width=""100%"">") & vbcrlf
				Response.Write("<tr><th></th><th>Extension</th><th>Count</th><th>%</th><th></th></tr>") & vbcrlf
				Response.Write("<input type=hidden name=data1 value=""Extension"">")
				Response.Write("<input type=hidden name=data2 value=""Count"">")
				Response.Write("<input type=hidden name=data3 value=""%"">")
				Response.Write("<input type=hidden name=cols value=3>")
			Else
				Response.Write("<p class=nodata>There is no data available for the report period selected.</p>") & vbcrlf
			End if
			Do While Not rsQuery.Eof

				If intTotal > 0 Then
					sngPercent = FormatPercent(rsQuery(1) / intTotal, 2)
				Else
					sngPercent = FormatPercent(0, 2)
				End If
			
				intRow = intRow + 1
				If (intRow Mod 2) = 1 Then
					strClass = "data"
				Else
					strClass = "dataalt"
				End If
			
				If intRow = 1 Then
					intMaxNumber = rsQuery(1)
				End If
			
				With Response
					.Write("<tr class=" & strClass & ">") & vbcrlf
					.Write("<td align=right>" & intRow & ".</td>")
					.Write("<td align=left>" & UCase(rsQuery(0)))
					.Write(DisplayCountry(rsQuery(0)) & "</td>")
					.Write("<td align=right>" & rsQuery(1) & "</td>")
					.Write("<td align=right>" & sngPercent & "</td>")
					.Write("<td align=left>")
					If blnShowGraph = True Then
						Call GenerateGraph(rsQuery(1), intMaxNumber, strClass)
					End If
					.Write("</td>")
					.Write("</tr>") & vbcrlf
					.Write("<input type=hidden name=data1 value=""" & rsQuery(0) & """>")
					.Write("<input type=hidden name=data2 value=" & rsQuery(1) & ">")
					.Write("<input type=hidden name=data3 value=""" & sngPercent & """>") & vbcrlf
				End With
				rsQuery.Movenext
			Loop
			
			If rsQuery.RecordCount > 0  Then
				With Response
					.Write("<tr class=total>")
					.Write("<td colspan=2 align=right>Total: </td>")
					.Write("<td align=right>" & intTotal & "</td>")
					.Write("<td align=right>&nbsp;</td>")
					.Write("<td align=right>&nbsp;</td>")
					.Write("</tr>")
					.Write("<input type=hidden name=total value=" & intTotal & ">")
					.Write("</table>") & vbcrlf
				End With
			End if
			
			rsQuery.Close : Set rsQuery = Nothing
			
			Call DisplayReportFooter()
			
		Case 19 : ' SEARCHES (SEARCH ENGINES)
			
			intTotal = CountSearches(datStart, datEnd)
			
			strSql = "SELECT " & SetItems(blnItems, intItems, True) & " n_value, " &_
				"COUNT(n_value) AS dc_count FROM (SELECT pl_s_id, k_site " &_
				"FROM " & strTablePrefix & "Keywords, " & strTablePrefix & "Referrers, " & strTablePrefix & "PageLog " &_
				"WHERE k_id = r_k_id " &_
				"AND r_id = pl_r_id " &_
				"AND pl_datetime BETWEEN " & FormatDatabaseDate(datStart & " 00:00:00") &_
				" AND " & FormatDatabaseDate(datEnd & " 23:59:59") & " " &_
				"AND k_site <> 0) dt_Sites LEFT JOIN " & strTablePrefix & "Names " &_
				"ON dt_Sites.k_site = " & strTablePrefix & "Names.n_id " &_
				"GROUP BY n_value "
				
			If strDatabaseType = "MYSQL" Then
				strSql = strSql & "ORDER BY dc_count DESC, n_value ASC " & SetItems(blnItems, intItems, False)
			Else
				strSql = strSql & "ORDER BY COUNT(n_value) DESC, n_value ASC"
			End If

			Set rsQuery = Server.CreateObject("ADODB.Recordset")
			
			If strDatabaseType = "MYSQL" Then
				rsQuery.CursorLocation = 3
			End If
			
			rsQuery.Open strSql, objConn, 3, 1, &H0000
			
			Call DisplayReportHeader(strReportGroup, strReportName, strDesc)
			
			If Not rsQuery.Eof Then
				Response.Write("<table border=0 cellpadding=4 cellspacing=0 width=""100%"">") & vbcrlf
				Response.Write("<tr><th></th><th>Site</th><th>Count</th><th>%</th><th></th></tr>") & vbcrlf
				Response.Write("<input type=hidden name=data1 value=""Site"">")
				Response.Write("<input type=hidden name=data2 value=""Count"">")
				Response.Write("<input type=hidden name=data3 value=""%"">")
				Response.Write("<input type=hidden name=cols value=3>")
			Else
				Response.Write("<p class=nodata>There is no data available for the report period selected.</p>") & vbcrlf
			End if
			Do While Not rsQuery.Eof

				If intTotal > 0 Then
					sngPercent = FormatPercent(rsQuery(1) / intTotal, 2)
				Else
					sngPercent = FormatPercent(0, 2)
				End If
			
				intRow = intRow + 1
				If (intRow Mod 2) = 1 Then
					strClass = "data"
				Else
					strClass = "dataalt"
				End If
			
				If intRow = 1 Then
					intMaxNumber = rsQuery(1)
				End If
			
				With Response
					.Write("<tr class=" & strClass & ">") & vbcrlf
					.Write("<td align=right>" & intRow & ".</td>")
					.Write("<td align=left>" & rsQuery(0) & "</td>")
					.Write("<td align=right>" & rsQuery(1) & "</td>")
					.Write("<td align=right>" & sngPercent & "</td>")
					.Write("<td align=left>")
					If blnShowGraph = True Then
						Call GenerateGraph(rsQuery(1), intMaxNumber, strClass)
					End If
					.Write("</td>")
					.Write("</tr>") & vbcrlf
					.Write("<input type=hidden name=data1 value=""" & FormatExportData(rsQuery(0)) & """>")
					.Write("<input type=hidden name=data2 value=" & rsQuery(1) & ">")
					.Write("<input type=hidden name=data3 value=""" & sngPercent & """>") & vbcrlf
				End With
				rsQuery.Movenext
			Loop
			
			If rsQuery.RecordCount > 0  Then
				With Response
					.Write("<tr class=total>")
					.Write("<td colspan=2 align=right>Total: </td>")
					.Write("<td align=right>" & intTotal & "</td>")
					.Write("<td align=right>&nbsp;</td>")
					.Write("<td align=right>&nbsp;</td>")
					.Write("</tr>")
					.Write("<input type=hidden name=total value=" & intTotal & ">")
					.Write("</table>") & vbcrlf
				End With
			End if
			
			rsQuery.Close : Set rsQuery = Nothing
			
			Call DisplayReportFooter()
			
		Case 20 : ' KEYWORDS (SEARCH ENGINES)
			
			intTotal = CountKeywords(datStart, datEnd)
			
			strSql = "SELECT " & SetItems(blnItems, intItems, True) & "k_value, COUNT(k_value) AS dc_count " &_
				"FROM (SELECT pl_s_id, k_value " &_
				"FROM " & strTablePrefix & "Keywords, " & strTablePrefix & "Referrers, " & strTablePrefix & "PageLog " &_
				"WHERE k_id = r_k_id " &_
				"AND r_id = pl_r_id " &_
				"AND pl_datetime BETWEEN " & FormatDatabaseDate(datStart & " 00:00:00") &_
				" AND " & FormatDatabaseDate(datEnd & " 23:59:59") & " " &_
				"AND k_value <> '') dtKeywords " &_
				"GROUP BY k_value "
				
			If strDatabaseType = "MYSQL" Then
				strSql = strSql & "ORDER BY dc_count DESC, k_value ASC " & SetItems(blnItems, intItems, False)
			Else
				strSql = strSql & "ORDER BY COUNT(k_value) DESC, k_value ASC"
			End If

			Set rsQuery = Server.CreateObject("ADODB.Recordset")
			
			If strDatabaseType = "MYSQL" Then
				rsQuery.CursorLocation = 3
			End If
			
			rsQuery.Open strSql, objConn, 3, 1, &H0000
			
			Call DisplayReportHeader(strReportGroup, strReportName, strDesc)
			
			If Not rsQuery.Eof Then
				Response.Write("<table border=0 cellpadding=4 cellspacing=0 width=""100%"">") & vbcrlf
				Response.Write("<tr><th></th><th>Keywords</th><th>Count</th><th>%</th><th></th></tr>") & vbcrlf
				Response.Write("<input type=hidden name=data1 value=""Keywords"">")
				Response.Write("<input type=hidden name=data2 value=""Count"">")
				Response.Write("<input type=hidden name=data3 value=""%"">")
				Response.Write("<input type=hidden name=cols value=3>")
			Else
				Response.Write("<p class=nodata>There is no data available for the report period selected.</p>") & vbcrlf
			End if
			Do While Not rsQuery.Eof

				If intTotal > 0 Then
					sngPercent = FormatPercent(rsQuery(1) / intTotal, 2)
				Else
					sngPercent = FormatPercent(0, 2)
				End If
			
				intRow = intRow + 1
				If (intRow Mod 2) = 1 Then
					strClass = "data"
				Else
					strClass = "dataalt"
				End If
			
				If intRow = 1 Then
					intMaxNumber = rsQuery(1)
				End If
			
				With Response
					.Write("<tr class=" & strClass & ">") & vbcrlf
					.Write("<td align=right>" & intRow & ".</td>")
					.Write("<td align=left>" & rsQuery(0) & "</td>")
					.Write("<td align=right>" & rsQuery(1) & "</td>")
					.Write("<td align=right>" & sngPercent & "</td>")
					.Write("<td align=left>")
					If blnShowGraph = True Then
						Call GenerateGraph(rsQuery(1), intMaxNumber, strClass)
					End If
					.Write("</td>")
					.Write("</tr>") & vbcrlf
					.Write("<input type=hidden name=data1 value=""" & FormatExportData(rsQuery(0)) & """>")
					.Write("<input type=hidden name=data2 value=" & rsQuery(1) & ">")
					.Write("<input type=hidden name=data3 value=""" & sngPercent & """>") & vbcrlf
				End With
				rsQuery.Movenext
			Loop
			
			If rsQuery.RecordCount > 0  Then
				With Response
					.Write("<tr class=total>")
					.Write("<td colspan=2 align=right>Total: </td>")
					.Write("<td align=right>" & intTotal & "</td>")
					.Write("<td align=right>&nbsp;</td>")
					.Write("<td align=right>&nbsp;</td>")
					.Write("</tr>")
					.Write("<input type=hidden name=total value=" & intTotal & ">")
					.Write("</table>") & vbcrlf
				End With
			End if
			
			rsQuery.Close : Set rsQuery = Nothing
			
			Call DisplayReportFooter()

		Case 21 : ' ROBOTS (SEARCH ENGINES)
			
			intTotal = CountRobotPageViews(datStart, datEnd)
			
			strSql = "SELECT " & SetItems(blnItems, intItems, True) & "(SELECT n_value FROM " & strTablePrefix & "Names WHERE n_id = rl_robot), " &_
				"COUNT(rl_robot) AS dc_count FROM " & strTablePrefix & "RobotLog " &_
				"WHERE rl_datetime BETWEEN " & FormatDatabaseDate(datStart & " 00:00:00") &_
				" AND " & FormatDatabaseDate(datEnd & " 23:59:59") & " " &_
				"GROUP BY rl_robot "
				
			If strDatabaseType = "MYSQL" Then
				strSql = strSql & "ORDER BY dc_count DESC, rl_robot ASC " & SetItems(blnItems, intItems, False)
			Else
				strSql = strSql & "ORDER BY COUNT(rl_robot) DESC, rl_robot ASC"
			End If

			Set rsQuery = Server.CreateObject("ADODB.Recordset")
			
			If strDatabaseType = "MYSQL" Then
				rsQuery.CursorLocation = 3
			End If
			
			rsQuery.Open strSql, objConn, 3, 1, &H0000
			
			Call DisplayReportHeader(strReportGroup, strReportName, strDesc)
			
			If Not rsQuery.Eof Then
				Response.Write("<table border=0 cellpadding=4 cellspacing=0 width=""100%"">") & vbcrlf
				Response.Write("<tr><th></th><th>Robot</th><th>Page Views</th><th>%</th><th></th></tr>") & vbcrlf
				Response.Write("<input type=hidden name=data1 value=""Robot"">")
				Response.Write("<input type=hidden name=data2 value=""Page Views"">")
				Response.Write("<input type=hidden name=data3 value=""%"">")
				Response.Write("<input type=hidden name=cols value=3>")
			Else
				Response.Write("<p class=nodata>There is no data available for the report period selected.</p>") & vbcrlf
			End if
			Do While Not rsQuery.Eof

				If intTotal > 0 Then
					sngPercent = FormatPercent(rsQuery(1) / intTotal, 2)
				Else
					sngPercent = FormatPercent(0, 2)
				End If
			
				intRow = intRow + 1
				If (intRow Mod 2) = 1 Then
					strClass = "data"
				Else
					strClass = "dataalt"
				End If
			
				If intRow = 1 Then
					intMaxNumber = rsQuery(1)
				End If
			
				With Response
					.Write("<tr class=" & strClass & ">") & vbcrlf
					.Write("<td align=right>" & intRow & ".</td>")
					.Write("<td align=left>" & rsQuery(0) & "</td>")
					.Write("<td align=right>" & rsQuery(1) & "</td>")
					.Write("<td align=right>" & sngPercent & "</td>")
					.Write("<td align=left>")
					If blnShowGraph = True Then
						Call GenerateGraph(rsQuery(1), intMaxNumber, strClass)
					End If
					.Write("</td>")
					.Write("</tr>") & vbcrlf
					.Write("<input type=hidden name=data1 value=""" & FormatExportData(rsQuery(0)) & """>")
					.Write("<input type=hidden name=data2 value=" & rsQuery(1) & ">")
					.Write("<input type=hidden name=data3 value=""" & sngPercent & """>") & vbcrlf
				End With
				rsQuery.Movenext
			Loop
			
			If rsQuery.RecordCount > 0  Then
				With Response
					.Write("<tr class=total>")
					.Write("<td colspan=2 align=right>Total: </td>")
					.Write("<td align=right>" & intTotal & "</td>")
					.Write("<td align=right>&nbsp;</td>")
					.Write("<td align=right>&nbsp;</td>")
					.Write("</tr>")
					.Write("</table>") & vbcrlf
				End With
			End if
			
			rsQuery.Close : Set rsQuery = Nothing
			
			Call DisplayReportFooter()
			
		Case 22 : ' COUNTRIES (DEMOGRAPHICS)
			
			intTotal = CountCountries("", datStart, datEnd)
			
			strSql = "SELECT " & SetItems(blnItems, intItems, True) & "s_country, COUNT(s_country) AS dc_count FROM (" &_
				"SELECT DISTINCT s_id, s_country " &_
				"FROM " & strTablePrefix & "Sessions, " & strTablePrefix & "PageLog " &_
				"WHERE pl_s_id = s_id " &_
				"AND pl_datetime BETWEEN " & FormatDatabaseDate(datStart & " 00:00:00") &_
				" AND " & FormatDatabaseDate(datEnd & " 23:59:59") & " " &_
				"AND s_country <> '') dtCountries " &_
				"GROUP BY s_country "
				
			If strDatabaseType = "MYSQL" Then
				strSql = strSql & "ORDER BY dc_count DESC, s_country ASC " & SetItems(blnItems, intItems, False)
			Else
				strSql = strSql & "ORDER BY COUNT(s_country) DESC, s_country ASC"
			End If

			Set rsQuery = Server.CreateObject("ADODB.Recordset")
			
			If strDatabaseType = "MYSQL" Then
				rsQuery.CursorLocation = 3
			End If
			
			rsQuery.Open strSql, objConn, 3, 1, &H0000
			
			Call DisplayReportHeader(strReportGroup, strReportName, strDesc)
			
			If Not rsQuery.Eof Then
				Response.Write("<table border=0 cellpadding=4 cellspacing=0 width=""100%"">") & vbcrlf
				Response.Write("<tr><th></th><th>Country</th><th>Count</th><th>%</th><th></th></tr>") & vbcrlf
				Response.Write("<input type=hidden name=data1 value=""Country"">")
				Response.Write("<input type=hidden name=data2 value=""Count"">")
				Response.Write("<input type=hidden name=data3 value=""%"">")
				Response.Write("<input type=hidden name=cols value=3>")
			Else
				Response.Write("<p class=nodata>There is no data available for the report period selected.</p>") & vbcrlf
			End if
			Do While Not rsQuery.Eof

				If intTotal > 0 Then
					sngPercent = FormatPercent(rsQuery(1) / intTotal, 2)
				Else
					sngPercent = FormatPercent(0, 2)
				End If
			
				intRow = intRow + 1
				If (intRow Mod 2) = 1 Then
					strClass = "data"
				Else
					strClass = "dataalt"
				End If
			
				If intRow = 1 Then
					intMaxNumber = rsQuery(1)
				End If
			
				With Response
					.Write("<tr class=" & strClass & ">") & vbcrlf
					.Write("<td align=right>" & intRow & ".</td>")
					.Write("<td align=left>" & ConvertCountryCode(rsQuery(0)) & "</td>")
					.Write("<td align=right>" & rsQuery(1) & "</td>")
					.Write("<td align=right>" & sngPercent & "</td>")
					.Write("<td align=left>")
					If blnShowGraph = True Then
						Call GenerateGraph(rsQuery(1), intMaxNumber, strClass)
					End If
					.Write("</td>")
					.Write("</tr>") & vbcrlf
					.Write("<input type=hidden name=data1 value=""" & FormatExportData(ConvertCountryCode(rsQuery(0))) & """>")
					.Write("<input type=hidden name=data2 value=" & rsQuery(1) & ">")
					.Write("<input type=hidden name=data3 value=""" & sngPercent & """>") & vbcrlf
				End With
				rsQuery.Movenext
			Loop
			
			If rsQuery.RecordCount > 0  Then
				With Response
					.Write("<tr class=total>")
					.Write("<td colspan=2 align=right>Total: </td>")
					.Write("<td align=right>" & intTotal & "</td>")
					.Write("<td align=right>&nbsp;</td>")
					.Write("<td align=right>&nbsp;</td>")
					.Write("</tr>")
					.Write("<input type=hidden name=total value=" & intTotal & ">")
					.Write("</table>") & vbcrlf
				End With
			End if
			
			rsQuery.Close : Set rsQuery = Nothing
			
			Call DisplayReportFooter()
			
		Case 23 : ' BROWSERS (DEMOGRAPHICS)
			
			intTotal = CountBrowsers(datStart, datEnd)
			
			strSql = "SELECT " & SetItems(blnItems, intItems, True) & "(SELECT n_value FROM " & strTablePrefix & "Names WHERE n_id = s_browser), " &_
				"COUNT(s_browser) AS dc_count " &_
				"FROM (SELECT DISTINCT s_id, s_browser " &_
				"FROM " & strTablePrefix & "Sessions, " & strTablePrefix & "PageLog " &_
				"WHERE pl_s_id = s_id " &_
				"AND pl_datetime BETWEEN " & FormatDatabaseDate(datStart & " 00:00:00") &_
				" AND " & FormatDatabaseDate(datEnd & " 23:59:59") & " " &_
				"AND s_browser <> 0) dtBrowsers " &_
				"GROUP BY s_browser "
				
			If strDatabaseType = "MYSQL" Then
				strSql = strSql & "ORDER BY dc_count DESC, s_browser ASC " & SetItems(blnItems, intItems, False)
			Else
				strSql = strSql & "ORDER BY COUNT(s_browser) DESC, s_browser ASC"
			End If

			Set rsQuery = Server.CreateObject("ADODB.Recordset")
			
			If strDatabaseType = "MYSQL" Then
				rsQuery.CursorLocation = 3
			End If
			
			rsQuery.Open strSql, objConn, 3, 1, &H0000
			
			Call DisplayReportHeader(strReportGroup, strReportName, strDesc)
			
			If Not rsQuery.Eof Then
				Response.Write("<table border=0 cellpadding=4 cellspacing=0 width=""100%"">") & vbcrlf
				Response.Write("<tr><th></th><th>Browser</th><th>Count</th><th>%</th><th></th></tr>") & vbcrlf
				Response.Write("<input type=hidden name=data1 value=""Browser"">")
				Response.Write("<input type=hidden name=data2 value=""Count"">")
				Response.Write("<input type=hidden name=data3 value=""%"">")
				Response.Write("<input type=hidden name=cols value=3>")
			Else
				Response.Write("<p class=nodata>There is no data available for the report period selected.</p>") & vbcrlf
			End if
			Do While Not rsQuery.Eof

				If intTotal > 0 Then
					sngPercent = FormatPercent(rsQuery(1) / intTotal, 2)
				Else
					sngPercent = FormatPercent(0, 2)
				End If
			
				intRow = intRow + 1
				If (intRow Mod 2) = 1 Then
					strClass = "data"
				Else
					strClass = "dataalt"
				End If
			
				If intRow = 1 Then
					intMaxNumber = rsQuery(1)
				End If
			
				With Response
					.Write("<tr class=" & strClass & ">") & vbcrlf
					.Write("<td align=right>" & intRow & ".</td>")
					.Write("<td align=left>" & rsQuery(0) & "</td>")
					.Write("<td align=right>" & rsQuery(1) & "</td>")
					.Write("<td align=right>" & sngPercent & "</td>")
					.Write("<td align=left>")
					If blnShowGraph = True Then
						Call GenerateGraph(rsQuery(1), intMaxNumber, strClass)
					End If
					.Write("</td>")
					.Write("</tr>") & vbcrlf
					.Write("<input type=hidden name=data1 value=""" & FormatExportData(rsQuery(0)) & """>")
					.Write("<input type=hidden name=data2 value=" & rsQuery(1) & ">")
					.Write("<input type=hidden name=data3 value=""" & sngPercent & """>") & vbcrlf
				End With
				rsQuery.Movenext
			Loop
			
			If rsQuery.RecordCount > 0  Then
				With Response
					.Write("<tr class=total>")
					.Write("<td colspan=2 align=right>Total: </td>")
					.Write("<td align=right>" & intTotal & "</td>")
					.Write("<td align=right>&nbsp;</td>")
					.Write("<td align=right>&nbsp;</td>")
					.Write("</tr>")
					.Write("<input type=hidden name=total value=" & intTotal & ">")
					.Write("</table>") & vbcrlf
				End With
			End if
			
			rsQuery.Close : Set rsQuery = Nothing
			
			Call DisplayReportFooter()
			
		Case 24 : ' OPERATING SYSTEMS (DEMOGRAPHICS)
			
			intTotal = CountOS(datStart, datEnd)
			
			strSql = "SELECT " & SetItems(blnItems, intItems, True) & "(SELECT n_value FROM " & strTablePrefix & "Names WHERE n_id = s_os), " &_ 
				"COUNT(s_os) AS dc_count FROM (" &_
				"SELECT DISTINCT s_id, s_os " &_
				"FROM " & strTablePrefix & "Sessions, " & strTablePrefix & "PageLog " &_
				"WHERE pl_s_id = s_id " &_
				"AND pl_datetime BETWEEN " & FormatDatabaseDate(datStart & " 00:00:00") &_
				" AND " & FormatDatabaseDate(datEnd & " 23:59:59") & " " &_
				"AND s_os <> 0) dtOS " &_
				"GROUP BY s_os "
				
			If strDatabaseType = "MYSQL" Then
				strSql = strSql & "ORDER BY dc_count DESC, s_os ASC " & SetItems(blnItems, intItems, False)
			Else
				strSql = strSql & "ORDER BY COUNT(s_os) DESC, s_os ASC"
			End If

			Set rsQuery = Server.CreateObject("ADODB.Recordset")
			
			If strDatabaseType = "MYSQL" Then
				rsQuery.CursorLocation = 3
			End If
			
			rsQuery.Open strSql, objConn, 3, 1, &H0000
			
			Call DisplayReportHeader(strReportGroup, strReportName, strDesc)
			
			If Not rsQuery.Eof Then
				Response.Write("<table border=0 cellpadding=4 cellspacing=0 width=""100%"">") & vbcrlf
				Response.Write("<tr><th></th><th>Operating System</th><th>Count</th><th>%</th><th></th></tr>") & vbcrlf
				Response.Write("<input type=hidden name=data1 value=""Operating Systems"">")
				Response.Write("<input type=hidden name=data2 value=""Count"">")
				Response.Write("<input type=hidden name=data3 value=""%"">")
				Response.Write("<input type=hidden name=cols value=3>")
			Else
				Response.Write("<p class=nodata>There is no data available for the report period selected.</p>") & vbcrlf
			End if
			Do While Not rsQuery.Eof

				If intTotal > 0 Then
					sngPercent = FormatPercent(rsQuery(1) / intTotal, 2)
				Else
					sngPercent = FormatPercent(0, 2)
				End If
			
				intRow = intRow + 1
				If (intRow Mod 2) = 1 Then
					strClass = "data"
				Else
					strClass = "dataalt"
				End If
			
				If intRow = 1 Then
					intMaxNumber = rsQuery(1)
				End If
			
				With Response
					.Write("<tr class=" & strClass & ">") & vbcrlf
					.Write("<td align=right>" & intRow & ".</td>")
					.Write("<td align=left>" & rsQuery(0) & "</td>")
					.Write("<td align=right>" & rsQuery(1) & "</td>")
					.Write("<td align=right>" & sngPercent & "</td>")
					.Write("<td align=left>")
					If blnShowGraph = True Then
						Call GenerateGraph(rsQuery(1), intMaxNumber, strClass)
					End If
					.Write("</td>")
					.Write("</tr>") & vbcrlf
					.Write("<input type=hidden name=data1 value=""" & FormatExportData(rsQuery(0)) & """>")
					.Write("<input type=hidden name=data2 value=" & rsQuery(1) & ">")
					.Write("<input type=hidden name=data3 value=""" & sngPercent & """>") & vbcrlf
				End With
				rsQuery.Movenext
			Loop
			
			If rsQuery.RecordCount > 0  Then
				With Response
					.Write("<tr class=total>")
					.Write("<td colspan=2 align=right>Total: </td>")
					.Write("<td align=right>" & intTotal & "</td>")
					.Write("<td align=right>&nbsp;</td>")
					.Write("<td align=right>&nbsp;</td>")
					.Write("</tr>")
					.Write("<input type=hidden name=total value=" & intTotal & ">")
					.Write("</table>") & vbcrlf
				End With
			End if
			
			rsQuery.Close : Set rsQuery = Nothing
			
			Call DisplayReportFooter()
			
		Case 25 : ' LANGUAGES (DEMOGRAPHICS)
			
			intTotal = CountLanguages(datStart, datEnd)

			strSql = "SELECT " & SetItems(blnItems, intItems, True) & "s_language, COUNT(s_language) AS dc_count FROM (" &_
				"SELECT DISTINCT s_id, s_language " &_
				"FROM " & strTablePrefix & "Sessions, " & strTablePrefix & "PageLog " &_
				"WHERE pl_s_id = s_id " &_
				"AND pl_datetime BETWEEN " & FormatDatabaseDate(datStart & " 00:00:00") &_
				" AND " & FormatDatabaseDate(datEnd & " 23:59:59") & " " &_
				"AND s_language <> '') dtLanguages " &_
				"GROUP BY s_language "
				
			If strDatabaseType = "MYSQL" Then
				strSql = strSql & "ORDER BY dc_count DESC, s_language ASC " & SetItems(blnItems, intItems, False)
			Else
				strSql = strSql & "ORDER BY COUNT(s_language) DESC, s_language ASC"
			End If

			Set rsQuery = Server.CreateObject("ADODB.Recordset")
			
			If strDatabaseType = "MYSQL" Then
				rsQuery.CursorLocation = 3
			End If
			
			rsQuery.Open strSql, objConn, 3, 1, &H0000
			
			Call DisplayReportHeader(strReportGroup, strReportName, strDesc)
			
			If Not rsQuery.Eof Then
				Response.Write("<table border=0 cellpadding=4 cellspacing=0 width=""100%"">") & vbcrlf
				Response.Write("<tr><th></th><th>Language</th><th>Count</th><th>%</th><th></th></tr>") & vbcrlf
				Response.Write("<input type=hidden name=data1 value=""Language"">")
				Response.Write("<input type=hidden name=data2 value=""Count"">")
				Response.Write("<input type=hidden name=data3 value=""%"">")
				Response.Write("<input type=hidden name=cols value=3>")
			Else
				Response.Write("<p class=nodata>There is no data available for the report period selected.</p>") & vbcrlf
			End if
			Do While Not rsQuery.Eof

				If intTotal > 0 Then
					sngPercent = FormatPercent(rsQuery(1) / intTotal, 2)
				Else
					sngPercent = FormatPercent(0, 2)
				End If
			
				intRow = intRow + 1
				If (intRow Mod 2) = 1 Then
					strClass = "data"
				Else
					strClass = "dataalt"
				End If
			
				If intRow = 1 Then
					intMaxNumber = rsQuery(1)
				End If
			
				With Response
					.Write("<tr class=" & strClass & ">") & vbcrlf
					.Write("<td align=right>" & intRow & ".</td>")
					.Write("<td align=left>" & ConvertToLanguageName(rsQuery(0)) & "</td>")
					.Write("<td align=right>" & rsQuery(1) & "</td>")
					.Write("<td align=right>" & sngPercent & "</td>")
					.Write("<td align=left>")
					If blnShowGraph = True Then
						Call GenerateGraph(rsQuery(1), intMaxNumber, strClass)
					End If
					.Write("</td>")
					.Write("</tr>") & vbcrlf
					.Write("<input type=hidden name=data1 value=""" & FormatExportData(ConvertToLanguageName(rsQuery(0))) & """>")
					.Write("<input type=hidden name=data2 value=" & rsQuery(1) & ">")
					.Write("<input type=hidden name=data3 value=""" & sngPercent & """>") & vbcrlf
				End With
				rsQuery.Movenext
			Loop
			
			If rsQuery.RecordCount > 0  Then
				With Response
					.Write("<tr class=total>")
					.Write("<td colspan=2 align=right>Total: </td>")
					.Write("<td align=right>" & intTotal & "</td>")
					.Write("<td align=right>&nbsp;</td>")
					.Write("<td align=right>&nbsp;</td>")
					.Write("</tr>")
					.Write("<input type=hidden name=total value=" & intTotal & ">")
					.Write("</table>") & vbcrlf
				End With
			End if
			
			rsQuery.Close : Set rsQuery = Nothing
			
			Call DisplayReportFooter()
			
		Case 26 : ' SCREEN AREA (DEMOGRAPHICS)
			
			intTotal = CountScreenAreas(datStart, datEnd)

			strSql = "SELECT " & SetItems(blnItems, intItems, True) & "(SELECT n_value FROM " & strTablePrefix & "Names WHERE n_id = s_screenarea), " &_
				"COUNT(s_screenarea) AS dc_count FROM (" &_
				"SELECT DISTINCT s_id, s_screenarea " &_
				"FROM " & strTablePrefix & "Sessions, " & strTablePrefix & "PageLog " &_
				"WHERE pl_s_id = s_id " &_
				"AND pl_datetime BETWEEN " & FormatDatabaseDate(datStart & " 00:00:00") &_
				" AND " & FormatDatabaseDate(datEnd & " 23:59:59") & " " &_
				"AND s_screenarea <> 0) dtScreenarea " &_
				"GROUP BY s_screenarea "
				
			If strDatabaseType = "MYSQL" Then
				strSql = strSql & "ORDER BY dc_count DESC, s_screenarea ASC " & SetItems(blnItems, intItems, False)
			Else
				strSql = strSql & "ORDER BY COUNT(s_screenarea) DESC, s_screenarea ASC"
			End If

			Set rsQuery = Server.CreateObject("ADODB.Recordset")
			
			If strDatabaseType = "MYSQL" Then
				rsQuery.CursorLocation = 3
			End If
			
			rsQuery.Open strSql, objConn, 3, 1, &H0000
			
			Call DisplayReportHeader(strReportGroup, strReportName, strDesc)
			
			If Not rsQuery.Eof Then
				Response.Write("<table border=0 cellpadding=4 cellspacing=0 width=""100%"">") & vbcrlf
				Response.Write("<tr><th></th><th>Screen Area</th><th>Count</th><th>%</th><th></th></tr>") & vbcrlf
				Response.Write("<input type=hidden name=data1 value=""Screen Area"">")
				Response.Write("<input type=hidden name=data2 value=""Count"">")
				Response.Write("<input type=hidden name=data3 value=""%"">")
				Response.Write("<input type=hidden name=cols value=3>")
			Else
				Response.Write("<p class=nodata>There is no data available for the report period selected.</p>") & vbcrlf
			End if
			Do While Not rsQuery.Eof

				If intTotal > 0 Then
					sngPercent = FormatPercent(rsQuery(1) / intTotal, 2)
				Else
					sngPercent = FormatPercent(0, 2)
				End If
			
				intRow = intRow + 1
				If (intRow Mod 2) = 1 Then
					strClass = "data"
				Else
					strClass = "dataalt"
				End If
			
				If intRow = 1 Then
					intMaxNumber = rsQuery(1)
				End If
			
				With Response
					.Write("<tr class=" & strClass & ">") & vbcrlf
					.Write("<td align=right>" & intRow & ".</td>")
					.Write("<td align=left>" & rsQuery(0) & "</td>")
					.Write("<td align=right>" & rsQuery(1) & "</td>")
					.Write("<td align=right>" & sngPercent & "</td>")
					.Write("<td align=left>")
					If blnShowGraph = True Then
						Call GenerateGraph(rsQuery(1), intMaxNumber, strClass)
					End If
					.Write("</td>")
					.Write("</tr>") & vbcrlf
					.Write("<input type=hidden name=data1 value=""" & rsQuery(0) & """>")
					.Write("<input type=hidden name=data2 value=" & rsQuery(1) & ">")
					.Write("<input type=hidden name=data3 value=""" & sngPercent & """>") & vbcrlf
				End With
				rsQuery.Movenext
			Loop
			
			If rsQuery.RecordCount > 0  Then
				With Response
					.Write("<tr class=total>")
					.Write("<td colspan=2 align=right>Total: </td>")
					.Write("<td align=right>" & intTotal & "</td>")
					.Write("<td align=right>&nbsp;</td>")
					.Write("<td align=right>&nbsp;</td>")
					.Write("</tr>")
					.Write("<input type=hidden name=total value=" & intTotal & ">")
					.Write("</table>") & vbcrlf
				End With
			End if
			
			rsQuery.Close : Set rsQuery = Nothing
			
			Call DisplayReportFooter()
			
		Case 27 : ' USER AGENTS (DEMOGRAPHICS)
			
			intTotal = CountUserAgents(datStart, datEnd)
			
			strSql = "SELECT " & SetItems(blnItems, intItems, True) & "(SELECT n_value FROM " & strTablePrefix & "Names WHERE n_id = s_useragent), " &_
				"COUNT(s_useragent) AS dc_count FROM (" &_
				"SELECT DISTINCT s_id, s_useragent " &_
				"FROM " & strTablePrefix & "Sessions, " & strTablePrefix & "PageLog " &_
				"WHERE pl_s_id = s_id " &_
				"AND pl_datetime BETWEEN " & FormatDatabaseDate(datStart & " 00:00:00") &_
				" AND " & FormatDatabaseDate(datEnd & " 23:59:59") & " " &_
				"AND s_useragent <> 0) dtUseragents " &_
				"GROUP BY s_useragent "
				
			If strDatabaseType = "MYSQL" Then
				strSql = strSql & "ORDER BY dc_count DESC, s_useragent ASC " & SetItems(blnItems, intItems, False)
			Else
				strSql = strSql & "ORDER BY COUNT(s_useragent) DESC, s_useragent ASC"
			End If

			Set rsQuery = Server.CreateObject("ADODB.Recordset")
			
			If strDatabaseType = "MYSQL" Then
				rsQuery.CursorLocation = 3
			End If
			
			rsQuery.Open strSql, objConn, 3, 1, &H0000
			
			Call DisplayReportHeader(strReportGroup, strReportName, strDesc)
			
			If Not rsQuery.Eof Then
				Response.Write("<table border=0 cellpadding=4 cellspacing=0 width=""100%"">") & vbcrlf
				Response.Write("<tr><th></th><th>User Agent</th><th>Count</th><th>%</th><th></th></tr>") & vbcrlf
				Response.Write("<input type=hidden name=data1 value=""User Agent"">")
				Response.Write("<input type=hidden name=data2 value=""Count"">")
				Response.Write("<input type=hidden name=data3 value=""%"">")
				Response.Write("<input type=hidden name=cols value=3>")
			Else
				Response.Write("<p class=nodata>There is no data available for the report period selected.</p>") & vbcrlf
			End if
			Do While Not rsQuery.Eof

				If intTotal > 0 Then
					sngPercent = FormatPercent(rsQuery(1) / intTotal, 2)
				Else
					sngPercent = FormatPercent(0, 2)
				End If
			
				intRow = intRow + 1
				If (intRow Mod 2) = 1 Then
					strClass = "data"
				Else
					strClass = "dataalt"
				End If
			
				If intRow = 1 Then
					intMaxNumber = rsQuery(1)
				End If
			
				With Response
					.Write("<tr class=" & strClass & ">") & vbcrlf
					.Write("<td align=right>" & intRow & ".</td>")
					.Write("<td align=left>" & rsQuery(0) & "</td>")
					.Write("<td align=right>" & rsQuery(1) & "</td>")
					.Write("<td align=right>" & sngPercent & "</td>")
					.Write("<td align=left>")
					If blnShowGraph = True Then
						Call GenerateGraph(rsQuery(1), intMaxNumber, strClass)
					End If
					.Write("</td>")
					.Write("</tr>") & vbcrlf
					.Write("<input type=hidden name=data1 value=""" & FormatExportData(rsQuery(0)) & """>")
					.Write("<input type=hidden name=data2 value=" & rsQuery(1) & ">")
					.Write("<input type=hidden name=data3 value=""" & sngPercent & """>") & vbcrlf
				End With
				rsQuery.Movenext
			Loop
			
			If rsQuery.RecordCount > 0  Then
				With Response
					.Write("<tr class=total>")
					.Write("<td colspan=2 align=right>Total: </td>")
					.Write("<td align=right>" & intTotal & "</td>")
					.Write("<td align=right>&nbsp;</td>")
					.Write("<td align=right>&nbsp;</td>")
					.Write("</tr>")
					.Write("<input type=hidden name=total value=" & intTotal & ">")
					.Write("</table>") & vbcrlf
				End With
			End if
			
			rsQuery.Close : Set rsQuery = Nothing
			
			Call DisplayReportFooter()
			
		End Select
		
	End Sub
	
	Public Sub DisplayDateChooser(datDate, strForm)
	
	With Response
		.Write("<input type=textbox name=" & strForm & "display ")
		.Write("value=""" & FormatDisplayDate(datDate, strShortDate) & """ size=12 readonly>")
		.Write("<input type=hidden name=" & strForm & "date value=""" & datDate & """>&nbsp;")
		.Write("<a href=""javascript:calendar('report." & strForm & "',document.report." & strForm & "date.value);"">")
		.Write("<img src=""images/calendar.gif"" border=0></a>")
	End With
	
	End Sub
	
	Public Sub DisplayItemsChooser(intReportItems)
		
		With Response
			.Write("<select name=items>")
			.Write("<option value=10")
			If intReportItems = 10 Then
				.Write(" selected")
			End If
			.Write(">10 Items</option>")
			.Write("<option value=20")
			If intReportItems = 20 Then
				.Write(" selected")
			End If
			.Write(">20 Items</option>")
			.Write("<option value=50")
			If intReportItems = 50 Then
				.Write(" selected")
			End If
			.Write(">50 Items</option>")
			.Write("<option value=100")
			If intReportItems = 100 Then
				.Write(" selected")
			End If
			.Write(">100 Items</option>")
			.Write("<option value=200")
			If intReportItems = 200 Then
				.Write(" selected")
			End If
			.Write(">200 Items</option>")
			.Write("<option value=500")
			If intReportItems = 500 Then
				.Write(" selected")
			End If
			.Write(">500 Items</option>")
			.Write("</select>")
		End With
		
	End Sub
	
	Public Sub DisplayReportList(intReport)
		
		Dim strClass, strLastGroup, intLoop
		Dim aryReports : aryReports = GetReportArray()
		Dim strGroup : strGroup = aryReports(intReport, 1)
		Dim intGroup : intGroup = 0

		For intLoop = 0 To UBound(aryReports)
			If CInt(aryReports(intLoop, 2)) = 1 Then
				If strLastGroup <> aryReports(intLoop, 1) Then
					
					intGroup = intGroup + 1
					
					If strGroup <> aryReports(intLoop, 1) Then
						strClass = "display: none;"
					Else
						strClass = ""
					End If
					
					If strLastGroup <> "" Then
						Response.Write "</td></tr></table>"
					End If

					With Response
						.Write("<tr>")
						.Write("<td width=12 align=center><img src=""images/lt_grey_arrow_dn.gif"" border=0></td>")
						.Write("<td colspan=2 width=""100%""><a href=""javascript:showhide('" & intGroup & "');"" class=chtitle>")
						.Write(aryReports(intLoop, 1) & "</a></td>")
						.Write("</tr>" & vbcrlf)
						.Write("<tr><td colspan=3>")
						.Write("<table cellpadding=0 cellspacing=0 border=0 id=" & intGroup & " width=""100%""")
						If strClass <> "" Then
							.Write (" style=""" & strClass & """>")
						Else
							.Write (">")
						End If
					End With
				End If
				If CInt(intReport) = CInt(intLoop) Then
					strClass = "chselected"
				Else
					strClass = "chitem"
				End If
				With Response
					.Write("<tr>")
					.Write("<td width=12></td>")
					.Write("<td width=12><img src=""images/lt_grey_arrow_sm.gif"" border=0></td>")
					.Write("<td><a href=""javascript:GenerateReport(" & intLoop & ");"" class=" & strClass & ">")
					.Write(aryReports(intLoop, 0) & "</a></td>")
					.Write("</tr>" & vbcrlf)
				End With
				strLastGroup = aryReports(intLoop, 1)
			End If
		Next
		
		Response.Write("</td></tr></table>")
		
	End Sub
	
	Private Sub DisplayReportHeader(strReportGroup, strReportName, strText)
	
		With Response
			.Write("<table border=0 cellpadding=0 cellspacing=0>")
			.Write("<tr><td><span class=name>")
			.Write(strReportGroup & "&nbsp;&raquo;&nbsp;" & strReportName & "</td><td align=right>")
			.Write("<a href=""javascript: exportreport('CSV')""><img src=""images/csv.gif"" alt=""CSV File"" border=0></a>")
			.Write("<a href=""javascript: printpreview();""><img src=""images/print.gif"" alt=""Print Preview"" border=0></a>")
			.Write("<a href=""javascript: showhelp('reports','" & FormatBookmark(strReportGroup, strReportName) & "');"">")
			.Write("<img src=""images/help.gif"" alt=""Help"" border=0></a>")
			.Write("</td></tr>")
			.Write("<tr><td colspan=2>")
			.Write("<table border=0 cellpadding=0 cellspacing=0 width=""100%"" class=report>")
			.Write("<tr><td class=header><table border=0 cellpadding=0 cellspacing=0>")
			.Write("<tr>")
			.Write("<td width=20><img src=""images/grey_arrow.gif""></td>")
			.Write("<td class=description>")
			.Write(strText & "</td>")
			.Write("</tr></table></td></tr>")
			.Write("<form name=exportform method=post>")
			.Write("<input type=hidden name=site value=""" & strSiteName & """>")
			.Write("<input type=hidden name=desc value=""" & strText & """>")
			.Write("<input type=hidden name=report value=""" & strReportGroup & " - " & strReportName & """>")
			.Write("<input type=hidden name=export value=""" & FormatDateTime(Now(), 0) & """>")
			.Write("<tr><td class=report>")
		End With
	
	End Sub
	
	Private Function FormatBookmark(strGroup, strReport)
	
		Dim strTemp
	
		strTemp = strGroup & strReport
		strTemp = Replace(strTemp, "'", "")
		strTemp = Replace(strTemp, " ", "")
		strTemp = Replace(strTemp, "&", "")
		
		FormatBookmark = strTemp
	
	End Function
	
	Private Sub DisplayReportFooter()
	
		With Response
			.Write ("</td></tr></form></table></td></tr></table>")
		End With
	
	End Sub
	
	Private Sub GenerateGraph(intNumber, intMaxNumber, strClass)
	
		Dim intWidth, intMaxWidth
		
		intMaxwidth = 150
		
		If intMaxNumber > 0 Then
			intWidth = Round(((intNumber / intMaxNumber) * intMaxWidth), 0)
			
			With Response
				.Write("<table cellpadding=0 cellspacing=0 border=0>")
				.Write("<tr>")
				.Write("<td class=graph" & strClass & " width=" & intWidth & ">")
				.Write("<img src=""images/spacer.gif"" width=" & intWidth & " height=15 alt=""" & intNumber & """>")
				.Write("</td>")
				.Write("</tr>")
				.Write("</table>")
			End With
		
		End If
	
	End Sub
	
	'************
	'* COUNTING *
	'************
	
	Public Function CountPageViews(strScriptName, datStart, datEnd)
	
		Dim intTemp

		strSql	= "SELECT COUNT(pl_pn_id) FROM " &_
			"(SELECT pl_pn_id " &_
			"FROM " & strTablePrefix & "PageLog, " & strTablePrefix & "PageNames " &_
			"WHERE pl_pn_id = pn_id " &_
			"AND pl_datetime BETWEEN " & FormatDatabaseDate(datStart & " 00:00:00") &_
			" AND " & FormatDatabaseDate(datEnd & " 23:59:59")
		If strScriptName <> "" Then
			strSql = strSql & " AND pn_page LIKE '" & strScriptName & "'"
		End If
		
		strSql = strSql & ") dtPageLog "
		
		Dim rsCount : Set rsCount = Server.CreateObject("ADODB.Recordset")
		rsCount.Open strSql, objConn, 3, 1, &H0000
		
		If Not rsCount.Eof Then
			intTemp = rsCount(0)
		Else
			intTemp = 0
		End If
		rsCount.Close : Set rsCount = Nothing
		
		CountPageViews = intTemp
		
	End Function
	
	Public Function CountFileTypes(strFileType, datStart, datEnd)
	
		Dim intTemp

		strSql	= "SELECT COUNT(pl_pn_id) FROM " &_
			"(SELECT pl_pn_id " &_
			"FROM " & strTablePrefix & "PageLog, " & strTablePrefix & "PageNames " &_
			"WHERE pl_pn_id = pn_id " &_
			"AND pl_datetime BETWEEN " & FormatDatabaseDate(datStart & " 00:00:00") &_
			" AND " & FormatDatabaseDate(datEnd & " 23:59:59")
		If strFileType <> "" Then
			strSql = strSql & " AND pn_extension LIKE '" & strScriptName & "'"
		End If
		
		strSql = strSql & "AND pn_extension <> '') dtPageLog "
		
		Dim rsCount : Set rsCount = Server.CreateObject("ADODB.Recordset")
		rsCount.Open strSql, objConn, 3, 1, &H0000
		
		If Not rsCount.Eof Then
			intTemp = rsCount(0)
		Else
			intTemp = 0
		End If
		rsCount.Close : Set rsCount = Nothing
		
		CountFileTypes = intTemp
		
	End Function
	
	Public Function CountReferrers(strType, datStart, datEnd)

		Dim intTemp

		Select Case strType	

		Case "Domain"

			strSql	= "SELECT COUNT(rn_domain) FROM " &_
				"(SELECT rn_domain " &_
				"FROM " & strTablePrefix & "PageLog, " & strTablePrefix & "Referrers, " & strTablePrefix & "ReferrerNames " &_
				"WHERE pl_r_id = r_id " &_
				"AND r_rn_id = rn_id " &_
				"AND pl_datetime BETWEEN " & FormatDatabaseDate(datStart & " 00:00:00") &_
				" AND " & FormatDatabaseDate(datEnd & " 23:59:59") & " "
				
				If blnSiteAliases = True Then
					strSql = strSql & "AND rn_host NOT IN (" & FormatSiteAliases(strSiteAliases) & ") "
				End If
				strSql = strSql & "AND pl_r_id <> 0 " &_
				"AND rn_domain <> '') dtPageLog "

		Case "Host"

			strSql	= "SELECT COUNT(rn_host) FROM " &_
				"(SELECT rn_host " &_
				"FROM " & strTablePrefix & "PageLog, " & strTablePrefix & "Referrers, " & strTablePrefix & "ReferrerNames " &_
				"WHERE pl_r_id = r_id " &_
				"AND r_rn_id = rn_id " &_
				"AND pl_datetime BETWEEN " & FormatDatabaseDate(datStart & " 00:00:00") &_
				" AND " & FormatDatabaseDate(datEnd & " 23:59:59") & " "
				
				If blnSiteAliases = True Then
					strSql = strSql & "AND rn_host NOT IN (" & FormatSiteAliases(strSiteAliases) & ") "
				End If
				strSql = strSql & "AND pl_r_id <> 0 " &_
				"AND rn_host <> '') dtPageLog "
				
		Case "Page"

			strSql	= "SELECT COUNT(rn_page) FROM " &_
				"(SELECT pl_datetime, rn_page " &_
				"FROM " & strTablePrefix & "PageLog, " & strTablePrefix & "Referrers, " & strTablePrefix & "ReferrerNames " &_
				"WHERE pl_r_id = r_id " &_
				"AND r_rn_id = rn_id " &_
				"AND pl_datetime BETWEEN " & FormatDatabaseDate(datStart & " 00:00:00") &_
				" AND " & FormatDatabaseDate(datEnd & " 23:59:59") & " "
				
				If blnSiteAliases = True Then
					strSql = strSql & "AND rn_host NOT IN (" & FormatSiteAliases(strSiteAliases) & ") "
				End If
				strSql = strSql & "AND pl_r_id <> 0 " &_
				"AND rn_page <> '') dtPageLog "

		Case "Extension"

			strSql	= "SELECT COUNT(rn_extension) FROM " &_
				"(SELECT pl_datetime, rn_extension " &_
				"FROM " & strTablePrefix & "PageLog, " & strTablePrefix & "Referrers, " & strTablePrefix & "ReferrerNames " &_
				"WHERE pl_r_id = r_id " &_
				"AND r_rn_id = rn_id " &_
				"AND pl_datetime BETWEEN " & FormatDatabaseDate(datStart & " 00:00:00") &_
				" AND " & FormatDatabaseDate(datEnd & " 23:59:59") & " "
				
				If blnSiteAliases = True Then
					strSql = strSql & "AND rn_host NOT IN (" & FormatSiteAliases(strSiteAliases) & ") "
				End If
				strSql = strSql & "AND pl_r_id <> 0 " &_
				"AND rn_extension <> '') dtPageLog "

		Case Else 'Url

			strSql	= "SELECT COUNT(r_url) FROM " &_
				"(SELECT pl_datetime, r_url " &_
				"FROM " & strTablePrefix & "PageLog, " & strTablePrefix & "Referrers, " & strTablePrefix & "ReferrerNames " &_
				"WHERE pl_r_id = r_id " &_
				"AND r_rn_id = rn_id " &_
				"AND pl_datetime BETWEEN " & FormatDatabaseDate(datStart & " 00:00:00") &_
				" AND " & FormatDatabaseDate(datEnd & " 23:59:59") & " "
				
				If blnSiteAliases = True Then
					strSql = strSql & "AND rn_host NOT IN (" & FormatSiteAliases(strSiteAliases) & ") "
				End If
				strSql = strSql & "AND pl_r_id <> 0 " &_
				"AND r_url <> '') dtPageLog "

		End Select

		Dim rsCount : Set rsCount = Server.CreateObject("ADODB.Recordset")
		rsCount.Open strSql, objConn, 3, 1, &H0000
		
		If Not rsCount.Eof Then
			intTemp = rsCount(0)
		Else
			intTemp = 0
		End If

		rsCount.Close : Set rsCount = Nothing

		CountReferrers = intTemp

	End Function
	
	Public Function CountUsersOnline()

		Dim intTemp
		
		Dim datNow : datNow = Now()

		strSql	= "SELECT COUNT(s_ip) " &_
			"FROM (SELECT DISTINCT s_ip FROM " & strTablePrefix & "PageLog, " & strTablePrefix & "Sessions WHERE pl_s_id = s_id " &_
			"AND pl_datetime BETWEEN " & FormatDatabaseDate(Dateadd("n", (0 - intSessionDuration), datNow)) &_
			" AND " & FormatDatabaseDate(datNow) & ") dtPageLog"

		Dim rsCount : Set rsCount = Server.CreateObject("ADODB.Recordset")
		rsCount.Open strSql, objConn, 3, 1, &H0000
		
		If Not rsCount.Eof Then
			intTemp = rsCount(0)
		Else
			intTemp = 0
		End If
		
		rsCount.Close : Set rsCount = Nothing

		CountUsersOnline = intTemp

	End Function
	
	Public Function CountVisits(datStart, datEnd)

		Dim intTemp
		
		strSql	= "SELECT COUNT(s_id) FROM " &_
			"(SELECT DISTINCT s_id FROM " & strTablePrefix & "Sessions, " & strTablePrefix & "PageLog " &_
			"WHERE pl_s_id = s_id " &_
			"AND pl_datetime BETWEEN " & FormatDatabaseDate(datStart & " 00:00:00") &_
			" AND " & FormatDatabaseDate(datEnd & " 23:59:59") & ") dtPageLog"

		Dim rsCount : Set rsCount = Server.CreateObject("ADODB.Recordset")
		rsCount.Open strSql, objConn, 3, 1, &H0000
		
		If Not rsCount.Eof Then
			intTemp = rsCount(0)
		Else
			intTemp = 0
		End If
		rsCount.Close : Set rsCount = Nothing
	
		CountVisits = intTemp
	
	End Function
	
	Public Function CountVisitors(datStart, datEnd)

		Dim intTemp

		strSql	= "SELECT COUNT(s_ip) FROM " &_
			"(SELECT DISTINCT s_ip FROM " & strTablePrefix & "Sessions, " & strTablePrefix & "PageLog " &_
			"WHERE pl_s_id = s_id " &_
			"AND pl_datetime BETWEEN " & FormatDatabaseDate(datStart & " 00:00:00") &_
			" AND " & FormatDatabaseDate(datEnd & " 23:59:59") & ") dtPageLog"

		Dim rsCount : Set rsCount = Server.CreateObject("ADODB.Recordset")
		rsCount.Open strSql, objConn, 3, 1, &H0000
		
		If Not rsCount.Eof Then
			intTemp = rsCount(0)
		Else
			intTemp = 0
		End If
		
		rsCount.Close : Set rsCount = Nothing
	
		CountVisitors = intTemp
	
	End Function
	
	Public Function CountDailyVisitors(datStart, datEnd)

		Dim intTemp

		strSql	= "SELECT SUM(Visitors) FROM (SELECT COUNT(s_ip) AS Visitors " &_
			"FROM (SELECT DISTINCT YEAR(pl_datetime) AS LogYear, " &_
			"MONTH(pl_datetime) AS LogMonth,  " &_
			"DAY(pl_datetime) AS LogDay, s_ip FROM (" &_
			"SELECT pl_datetime, s_ip FROM " & strTablePrefix & "PageLog, " & strTablePrefix & "Sessions " &_
			"WHERE pl_s_id = s_id) dtPageLog " &_
			"WHERE pl_datetime BETWEEN " & FormatDatabaseDate(datStart & " 00:00:00") &_
			" AND " & FormatDatabaseDate(datEnd & " 23:59:59") &_
			") dtSessions GROUP BY LogYear, LogMonth, LogDay " &_
			") dtSum"

		Dim rsCount : Set rsCount = Server.CreateObject("ADODB.Recordset")
		rsCount.Open strSql, objConn, 3, 1, &H0000
		
		If Not rsCount.Eof Then
			intTemp = rsCount(0)
		Else
			intTemp = 0
		End If
		
		If Not IsNumeric(intTemp) Then
			intTemp = 0
		End If
		
		rsCount.Close : Set rsCount = Nothing
	
		CountDailyVisitors = intTemp
	
	End Function
	
	Public Function CountDailyVisits(datStart, datEnd)

		Dim intTemp

		strSql	= "SELECT SUM(Visits) FROM (SELECT COUNT(pl_s_id) AS Visits " &_
			"FROM (SELECT DISTINCT YEAR(pl_datetime) AS LogYear, " &_
			"MONTH(pl_datetime) AS LogMonth,  DAY(pl_datetime) AS LogDay, pl_s_id " &_
			"FROM " & strTablePrefix & "PageLog " &_
			"WHERE pl_datetime BETWEEN " & FormatDatabaseDate(datStart & " 00:00:00") &_
			" AND " & FormatDatabaseDate(datEnd & " 23:59:59") &_
			") dtPageLog GROUP BY LogYear, LogMonth, LogDay " &_
			") dtSum"

		Dim rsCount : Set rsCount = Server.CreateObject("ADODB.Recordset")
		rsCount.Open strSql, objConn, 3, 1, &H0000
		
		If Not rsCount.Eof Then
			intTemp = rsCount(0)
		Else
			intTemp = 0
		End If
		
		If Not IsNumeric(intTemp) Then
			intTemp = 0
		End If
		
		rsCount.Close : Set rsCount = Nothing
	
		CountDailyVisits = intTemp
	
	End Function
	
	Public Function CountSearches(datStart, datEnd)

		Dim intTemp
	
		strSql = "SELECT COUNT(k_site) " &_
			"FROM (SELECT pl_s_id, k_site " &_
			"FROM " & strTablePrefix & "Keywords, " & strTablePrefix & "Referrers, " & strTablePrefix & "PageLog " &_
			"WHERE k_id = r_k_id " &_
			"AND r_id = pl_r_id " &_
			"AND pl_datetime BETWEEN " & FormatDatabaseDate(datStart & " 00:00:00") &_
			" AND " & FormatDatabaseDate(datEnd & " 23:59:59") & " " &_
			"AND k_site <> 0) dtSites "

		Dim rsCount : Set rsCount = Server.CreateObject("ADODB.Recordset")
		rsCount.Open strSql, objConn, 3, 1, &H0000
		
		If Not rsCount.Eof Then
			intTemp = rsCount(0)
		Else
			intTemp = 0
		End If
		
		rsCount.Close : Set rsCount = Nothing
	
		CountSearches = intTemp
	
	End Function
	
	Public Function CountKeywords(datStart, datEnd)

		Dim intTemp
	
			strSql = "SELECT COUNT(k_value) " &_
				"FROM (SELECT pl_s_id, k_value FROM " & strTablePrefix & "Keywords, " & strTablePrefix & "Referrers, " & strTablePrefix & "PageLog " &_
				"WHERE k_id = r_k_id " &_
				"AND r_id = pl_r_id " &_
				"AND pl_datetime BETWEEN " & FormatDatabaseDate(datStart & " 00:00:00") &_
				" AND " & FormatDatabaseDate(datEnd & " 23:59:59") & " " &_
				"AND k_value <> '') dtKeywords"

		Dim rsCount : Set rsCount = Server.CreateObject("ADODB.Recordset")
		rsCount.Open strSql, objConn, 3, 1, &H0000
		
		If Not rsCount.Eof Then
			intTemp = rsCount(0)
		Else
			intTemp = 0
		End If
		
		rsCount.Close : Set rsCount = Nothing
	
		CountKeywords = intTemp
	
	End Function
	
	Public Function CountRobotPageViews(datStart, datEnd)

		Dim intTemp
	
		strSql = "SELECT COUNT(*) FROM " & strTablePrefix & "RobotLog " &_
			"WHERE rl_datetime BETWEEN " & FormatDatabaseDate(datStart & " 00:00:00") &_
			" AND " & FormatDatabaseDate(datEnd & " 23:59:59")

		Dim rsCount : Set rsCount = Server.CreateObject("ADODB.Recordset")
		rsCount.Open strSql, objConn, 3, 1, &H0000
		
		If Not rsCount.Eof Then
			intTemp = rsCount(0)
		Else
			intTemp = 0
		End If
		
		rsCount.Close : Set rsCount = Nothing
	
		CountRobotPageViews = intTemp
	
	End Function
	
	Public Function CountCountries(strCode, datStart, datEnd)

		Dim intTemp
	
			strSql = "SELECT COUNT(s_country) FROM (" &_
				"SELECT DISTINCT s_id, s_country " &_
				"FROM " & strTablePrefix & "Sessions, " & strTablePrefix & "PageLog " &_
				"WHERE pl_s_id = s_id " &_
				"AND pl_datetime BETWEEN " & FormatDatabaseDate(datStart & " 00:00:00") &_
				" AND " & FormatDatabaseDate(datEnd & " 23:59:59") & " "
				
			If strCode <> "" Then
				strSql = strSql & "AND s_country = '" & strCode & "' "
			End If
				
			strSql = strSql & "AND s_country <> '') dtCountries"

		Dim rsCount : Set rsCount = Server.CreateObject("ADODB.Recordset")
		rsCount.Open strSql, objConn, 3, 1, &H0000
		
		If Not rsCount.Eof Then
			intTemp = rsCount(0)
		Else
			intTemp = 0
		End If
		
		rsCount.Close : Set rsCount = Nothing
	
		CountCountries = intTemp
	
	End Function
	
	Public Function CountBrowsers(datStart, datEnd)

		Dim intTemp
	
			strSql = "SELECT COUNT(s_browser) FROM (" &_
				"SELECT DISTINCT s_id, s_browser " &_
				"FROM " & strTablePrefix & "Sessions, " & strTablePrefix & "PageLog " &_
				"WHERE pl_s_id = s_id " &_
				"AND pl_datetime BETWEEN " & FormatDatabaseDate(datStart & " 00:00:00") &_
				" AND " & FormatDatabaseDate(datEnd & " 23:59:59") & " " &_
				"AND s_browser <> 0) dtBrowsers"

		Dim rsCount : Set rsCount = Server.CreateObject("ADODB.Recordset")
		rsCount.Open strSql, objConn, 3, 1, &H0000
		
		If Not rsCount.Eof Then
			intTemp = rsCount(0)
		Else
			intTemp = 0
		End If
		
		rsCount.Close : Set rsCount = Nothing
	
		CountBrowsers = intTemp
	
	End Function
	
	Public Function CountOS(datStart, datEnd)

		Dim intTemp
	
			strSql = "SELECT COUNT(s_os) FROM (" &_
				"SELECT DISTINCT s_id, s_os " &_
				"FROM " & strTablePrefix & "Sessions, " & strTablePrefix & "PageLog " &_
				"WHERE pl_s_id = s_id " &_
				"AND pl_datetime BETWEEN " & FormatDatabaseDate(datStart & " 00:00:00") &_
				" AND " & FormatDatabaseDate(datEnd & " 23:59:59") & " " &_
				"AND s_os <> 0) dtOS"

		Dim rsCount : Set rsCount = Server.CreateObject("ADODB.Recordset")
		rsCount.Open strSql, objConn, 3, 1, &H0000
		
		If Not rsCount.Eof Then
			intTemp = rsCount(0)
		Else
			intTemp = 0
		End If
		
		rsCount.Close : Set rsCount = Nothing
	
		CountOS = intTemp
	
	End Function
	
	Public Function CountLanguages(datStart, datEnd)

		Dim intTemp
	
			strSql = "SELECT COUNT(s_language) FROM (" &_
				"SELECT DISTINCT s_id, s_language " &_
				"FROM " & strTablePrefix & "Sessions, " & strTablePrefix & "PageLog " &_
				"WHERE pl_s_id = s_id " &_
				"AND pl_datetime BETWEEN " & FormatDatabaseDate(datStart & " 00:00:00") &_
				" AND " & FormatDatabaseDate(datEnd & " 23:59:59") & " " &_
				"AND s_language <> '') dtLanguages"

		Dim rsCount : Set rsCount = Server.CreateObject("ADODB.Recordset")
		rsCount.Open strSql, objConn, 3, 1, &H0000
		
		If Not rsCount.Eof Then
			intTemp = rsCount(0)
		Else
			intTemp = 0
		End If
		
		rsCount.Close : Set rsCount = Nothing
	
		CountLanguages = intTemp
	
	End Function
	
	Public Function CountScreenAreas(datStart, datEnd)

		Dim intTemp
	
			strSql = "SELECT COUNT(s_screenarea) FROM (" &_
				"SELECT DISTINCT s_id, s_screenarea " &_
				"FROM " & strTablePrefix & "Sessions, " & strTablePrefix & "PageLog " &_
				"WHERE pl_s_id = s_id " &_
				"AND pl_datetime BETWEEN " & FormatDatabaseDate(datStart & " 00:00:00") &_
				" AND " & FormatDatabaseDate(datEnd & " 23:59:59") & " " &_
				"AND s_screenarea <> 0) dtScreenareas"

		Dim rsCount : Set rsCount = Server.CreateObject("ADODB.Recordset")
		rsCount.Open strSql, objConn, 3, 1, &H0000
		
		If Not rsCount.Eof Then
			intTemp = rsCount(0)
		Else
			intTemp = 0
		End If
		
		rsCount.Close : Set rsCount = Nothing
	
		CountScreenAreas = intTemp
	
	End Function
	
	Public Function CountUserAgents(datStart, datEnd)

		Dim intTemp
	
			strSql = "SELECT COUNT(s_useragent) FROM (" &_
				"SELECT DISTINCT s_id, s_useragent " &_
				"FROM " & strTablePrefix & "Sessions, " & strTablePrefix & "PageLog " &_
				"WHERE pl_s_id = s_id " &_
				"AND pl_datetime BETWEEN " & FormatDatabaseDate(datStart & " 00:00:00") &_
				" AND " & FormatDatabaseDate(datEnd & " 23:59:59") & " " &_
				"AND s_useragent <> 0) dtScreenareas"

		Dim rsCount : Set rsCount = Server.CreateObject("ADODB.Recordset")
		rsCount.Open strSql, objConn, 3, 1, &H0000
		
		If Not rsCount.Eof Then
			intTemp = rsCount(0)
		Else
			intTemp = 0
		End If
		
		rsCount.Close : Set rsCount = Nothing
	
		CountUserAgents = intTemp
	
	End Function
	
	Public Function CountDistinctPages()
	
		Dim intTemp
	
		' COUNT DISTINCT PAGES
		strSql	= "SELECT COUNT(pn_page) FROM " &_
			"(SELECT DISTINCT pn_page FROM " & strTablePrefix & "PageNames) dtPages"

		Dim rsCount : Set rsCount = Server.CreateObject("ADODB.Recordset")
		rsCount.Open strSql, objConn, 3, 1, &H0000
		
		If Not rsCount.Eof Then
			intTemp = rsCount(0)
		Else
			intTemp = 0
		End If
		
		rsCount.Close : Set rsCount = Nothing
		
		CountDistinctPages = intTemp
		
	End Function
	
	Public Function CountDistinctUrls()
	
		Dim intTemp
	
		' COUNT DISTINCT URLS
		strSql	= "SELECT COUNT(pn_url) FROM " & strTablePrefix & "PageNames"

		Dim rsCount : Set rsCount = Server.CreateObject("ADODB.Recordset")
		rsCount.Open strSql, objConn, 3, 1, &H0000
		
		If Not rsCount.Eof Then
			intTemp = rsCount(0)
		Else
			intTemp = 0
		End If
		
		rsCount.Close : Set rsCount = Nothing
		
		CountDistinctUrls = intTemp
		
	End Function
	
	Public Function GetStartDate()
	
		Dim datTemp
		
		strSql = "SELECT MIN(pl_datetime) FROM " & strTablePrefix & "PageLog"

		Dim rsCount : Set rsCount = Server.CreateObject("ADODB.Recordset")
		rsCount.Open strSql, objConn, 3, 1, &H0000
		
		If Not rsCount.Eof Then
			datTemp = rsCount(0)
		Else
			datTemp = Empty
		End If
		
		rsCount.Close : Set rsCount = Nothing
		
		GetStartDate = datTemp
		
	End Function
	
	Private Function GetMaxPageViews(strType)
	
		Dim intTemp	
		
		If strType = "HOURLY" Then
		
			strSql	= "SELECT MAX(PageCount) FROM " &_
				"(SELECT COUNT(pl_pn_id) AS PageCount " &_
				"FROM " & strTablePrefix & "PageLog " &_
				"WHERE pl_datetime BETWEEN " & FormatDatabaseDate(datStart & " 00:00:00")&_
				" AND " & FormatDatabaseDate(datEnd & " 23:59:59") & " "
			If strDatabaseType = "MSSQL" Then
				strSql = strSql & "GROUP BY DATEPART(hh, pl_datetime)) dtPageLog"
			Else
				strSql = strSql & "GROUP BY HOUR(pl_datetime)) dtPageLog"
			End If
		Else
		
			strSql	= "SELECT MAX(PageCount) FROM " &_
				"(SELECT COUNT(pl_pn_id) AS PageCount " &_
				"FROM " & strTablePrefix & "PageLog " &_
				"WHERE pl_datetime BETWEEN " & FormatDatabaseDate(datStart & " 00:00:00")&_
				" AND " & FormatDatabaseDate(datEnd & " 23:59:59") & " " &_	
				"GROUP BY YEAR(pl_datetime), MONTH(pl_datetime), DAY(pl_datetime)) dtPageLog "
		
		End If
		
		Dim rsCount : Set rsCount = Server.CreateObject("ADODB.Recordset")
		rsCount.Open strSql, objConn, 3, 1, &H0000
		
		If Not rsCount.Eof Then
			intTemp = rsCount(0)
		Else
			intTemp = 0
		End If
		
		rsCount.Close
		Set rsCount = Nothing
	
		GetMaxPageViews = intTemp
	
	End Function
	
	Private Function GetMaxVisitors()
	
		Dim intTemp	
	
		strSql	= "SELECT MAX(VisitorCount) FROM " &_
				"(SELECT COUNT(s_ip) As VisitorCount FROM  " &_
				"(SELECT DISTINCT YEAR(pl_datetime) AS LogYear, MONTH(pl_datetime) AS LogMonth, " &_
				"DAY(pl_datetime) AS LogDay, s_ip FROM " & strTablePrefix & "PageLog, " & strTablePrefix & "Sessions " &_
				"WHERE pl_s_id = s_id " &_
				"AND pl_datetime BETWEEN " & FormatDatabaseDate(datStart & " 00:00:00") &_
				" AND " & FormatDatabaseDate(datEnd & " 23:59:59") & " ) dtSessions " &_
				"GROUP BY LogYear, LogMonth, LogDay) dtPageLog "

		Dim rsCount : Set rsCount = Server.CreateObject("ADODB.Recordset")
		rsCount.Open strSql, objConn, 3, 1, &H0000
		
		If Not rsCount.Eof Then
			intTemp = rsCount(0)
		Else
			intTemp = 0
		End If
		
		rsCount.Close
		Set rsCount = Nothing
	
		GetMaxVisitors = intTemp
	
	End Function
	
	Private Function GetMaxVisits()
	
		Dim intTemp	
	
		strSql	= "SELECT MAX(VisitCount) FROM " &_
				"(SELECT COUNT(pl_s_id) AS VisitCount " &_
				"FROM (SELECT DISTINCT YEAR(pl_datetime) AS LogYear, " &_
				"MONTH(pl_datetime) AS LogMonth,  DAY(pl_datetime) AS LogDay, pl_s_id " &_
				"FROM " & strTablePrefix & "PageLog " &_
				"WHERE pl_datetime BETWEEN " & FormatDatabaseDate(datStart & " 00:00:00") &_
				" AND " & FormatDatabaseDate(datEnd & " 23:59:59") &_
				") dtPageLog GROUP BY LogYear, LogMonth, LogDay) dtMaxPageLog "
		
		Dim rsCount : Set rsCount = Server.CreateObject("ADODB.Recordset")
		rsCount.Open strSql, objConn, 3, 1, &H0000
		
		If Not rsCount.Eof Then
			intTemp = rsCount(0)
		Else
			intTemp = 0
		End If
		
		rsCount.Close
		Set rsCount = Nothing
	
		GetMaxVisits = intTemp
	
	End Function

	Private Function FormatSiteAliases(strSiteAliases)
	
		' FORMAT SITE ALIASES
		If Len(strSiteAliases) > 0 Then
			
			' REMOVE SOME ILLEGAL CHARACTERS
			strSiteAliases = Replace(strSiteAliases, " ", "")
			strSiteAliases = Replace(strSiteAliases, "'", "")
			
			' CREATE ARRAY
			Dim arySiteAliases : arySiteAliases = Split(strSiteAliases, ",")
			strSiteAliases = ""
			
			Dim intLoop : For intLoop = 0 To UBound(arySiteAliases)
				strSiteAliases = strSiteAliases & "'" & arySiteAliases(intLoop) & "',"
			Next
			
			strSiteAliases = Mid(strSiteAliases, 1, Len(strSiteAliases) - 1)
		End If
		
		FormatSiteAliases = strSiteAliases
		
	End Function
	
	Private Function FormatLink(strUrl, strClass)
	
		Dim strTemp
	
		If Left(strSiteUrl, 4) = "http" Then
			strTemp = "<a href=""" & strSiteUrl & strUrl & """ target=""_new"" class=""" & strClass & """>" & strUrl & "</a>"
		Else
			strTemp = strUrl
		End If
		
		FormatLink = strTemp
	
	End Function
	
	Private Function DisplayCountry(strExtension)
		
		Dim strTemp
		
		If Left(Right(strExtension, 3), 1) = "." Then
			strTemp = ConvertCountryCode(UCase(Right(strExtension, 2)))
		End If
		
		If Len(strTemp) > 2 Then
			strTemp = " (" & strTemp & ")"
		End If
		
		DisplayCountry = strTemp
	
	End Function
	
	Private Function FormatExportData(strData)
	
		If Not IsNull(strData) Then
			strData = Replace(strData, ",", "%2C")
			strData = Replace(strData, """", "%22")
		End If
		
		FormatExportData = strData
	
	End Function
	
	Private Function TruncateUrl(strUrl)
	
		Dim strTemp, strBase, colItem, intLoop, intLength
		Dim objSearch, objBaseSearch, objBaseResults
		Dim strMatch, strStart, strEnd, intPosition
		Dim aryUrl, strQuerystring, strScriptName

		intLength	= 60
		strTemp		= strUrl

		If blnTruncateUrl = True Then
		
			If Len(strTemp) > intLength Then
	
				Set objBaseSearch	= New RegExp
				With objBaseSearch
					.Pattern 	= "(http|https)://[\w|\-|\.|:]+/"
					.IgnoreCase	= True
					.Global 	= False
				End With
	
				' CHECK FOR BASE URL			
				Set objBaseResults = objBaseSearch.Execute(strTemp)
	
				If objBaseResults.Count > 0 Then
	
					For Each colItem In objBaseResults
						strBase = colItem.Value	
					Next
	
					' REMOVE URL BASE BEFORE WE START LOOP
					strTemp = Mid(strTemp, Len(strBase) + 1)
				
					If InStr(strTemp, "?") > 0 Then
						aryUrl = Split(strTemp, "?")
						strScriptName = aryUrl(0)
						strQuerystring = aryUrl(1)
					Else
						strScriptName = strTemp
					End If
	
					For intLoop = 1 To 10
	
						Set objSearch		= New RegExp
						With objSearch
							.Pattern 	= "[\w|\.|\-|%|~]+/"
							.IgnoreCase	= True
							.Global 	= False
						End With
	
						' REPLACE PATTERN		
						strScriptName = objSearch.Replace(strScriptName, "##/")
	
						Set objSearch = Nothing
	
						If Len(strScriptName) + Len(strBase) + Len(strQuerystring) + 1 < intLength Then
							Exit For
						End If
	
					Next
	
					strScriptName = Replace(strScriptName, "##/", "../")
					
					' REASSEMBLE URL WITH BASE
					strTemp = strBase & strScriptName
					If Len(strQuerystring) > 0 Then
						strTemp = strTemp & "?" & strQuerystring
					End If

				End If
				
				If Len(strTemp) > intLength Then
					strTemp = Left(strTemp, intLength)
				End If
				
			End If
	
			Set objBaseSearch = Nothing
			Set objBaseResults = Nothing
		
		End If
		
		TruncateUrl = strTemp

	End Function
	
	Private Function SetItems(blnItems, intItems, blnTop)
	
		Dim strTemp
		
		If blnItems = True Then
			If strDatabaseType = "MYSQL" Then
				If blnTop = False Then
					strTemp = " LIMIT " & intItems
				End If
			Else
				If blnTop = True Then
					strTemp = "TOP " & intItems & " "
				End If
			End If
		End If
		
		SetItems = strTemp
	
	End Function
	
	Private Sub CheckVersion(strCurrent)
	
		Dim strSql : strSql = "SELECT c_value " &_
			"FROM " & strTablePrefix & "Config " &_
			"WHERE c_name = 'FJstats_Version'"
		
		Dim rsVersion : Set rsVersion = Server.CreateObject("ADODB.Recordset")
	
		rsVersion.Open strSql, objConn, 1, 2, &H0001
	
		If Not rsVersion.Eof Then
			If rsVersion(0) <> strCurrent Then
				Dim aryVersion : aryVersion = Split(rsVersion(0), " ")
				If UBound(aryVersion) = 1 And IsNumeric(aryVersion(0)) Then
					If CSng(aryVersion(0)) < 2.20 Then
						Call InsertConfig("Time_Offset", "0", "Reports", 2, 12, "select||int||-23,-22,-21,-20,-19,-18,-17,-16,-15,-14,-13,-12,-11,-10,-9,-8,-7,-6,-5,-4,-3,-2,-1,0,1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23||||")
					End If
				End If
				rsVersion(0) = strCurrent
				rsVersion.Update
			End If
		End If
	
		rsVersion.Close : Set rsVersion = Nothing
		
	End Sub

	Sub InsertConfig(strName, strValue, strGroup, intType, intOrder, strExtra)

		strSql = "INSERT INTO " & strTablePrefix & "Config (c_name, c_value, c_group, c_type, c_order, c_extra) VALUES" &_
			"(" & FormatDatabaseString(strName, 255) & ", " &_
			FormatDatabaseString(strValue, 255) & ", " &_
			FormatDatabaseString(strGroup, 255) & ", " &_
			intType & ", " &_
			intOrder & ", " &_
			FormatDatabaseString(strExtra, 255) & ")"
		objConn.Execute(strSql)

	End Sub	

	Private Function ConvertIPNumberToAddress(intIPNumber)

		Dim intTemp : intTemp = Int(intIPNumber)
		Dim aryIP(3)
		
		intTemp = intTemp + 2147483647
		
		aryIP(0) = Int(intTemp / 16777216) Mod 256
		intTemp = intTemp - 16777216 * aryIP(0)
		aryIP(1) = Int(intTemp / 65536) Mod 256
		intTemp = intTemp - 65536 * aryIP(1)
		aryIP(2) = Int(intTemp / 256) Mod 256
		intTemp = intTemp - 256 * aryIP(2)
		aryIP(3) = intTemp Mod 256
		
		Dim strTemp : strTemp = aryIP(0) & "." & aryIP(1) & "." & aryIP(2) & "." & aryIP(3)
		
		ConvertIPNumberToAddress = strTemp

	End Function
	
	Private Function ChooseIPWhois(strCountry)

		Dim strTemp
			
		Select Case strCountry

		' APNIC
		CASE "AF", "AS", "AU", "BD", "BN", "BT", "CC", "CK", "CN", "CX", _
		"FJ", "FM", "GU", "HK", "ID", "IN", "IO", "JP", "KH", "KI", _
		"KM", "KP", "KR", "LA", "LK", "MH", "MM", "MN", "MO", "MP", _
		"MV", "MY", "NC", "NF", "NP", "NR", "NU", "NZ", "PF", "PG", _
		"PH", "PK", "PN", "PW", "RE", "SB", "SG", "TF", "TH", "TK", _
		"TO", "TP", "TV", "TW", "VN", "VU", "WF", "WS", "YT"		
		
			strTemp = "APNIC"
		
		CASE "AG", "AI", "AQ", "BB", "BM", "BS", "BV", "CA", "CG", "DM", _
		"GD", "GP", "HM", "JM", "KN", "KY", "LC", "MQ", "PM", "PR", _
		"SH", "TC", "UM", "US", "VC", "VG", "VI"
		
			strTemp = "ARIN"

		CASE "AN", "AR", "AW", "BO", "BR", "BZ", "CL", "CO", "CR", "CU", _
		"DO", "EC", "FK", "GF", "GS", "GT", "GY", "HN", "HT", "MX", _
		"NI", "PA", "PE", "PY", "SR", "SV", "TT", "UY", "VE"
		
			strTemp = "LACNIC"

		CASE "AD", "AE", "AL", "AM", "AT", "AZ", "BA", "BE", "BG", _
		"BH", "BY", "CH", "CV", "CZ", "DE", "DK", "EE", "EH", _
		"ES", "FI", "FO", "FR", "GE", "GI", "GL", "GN", "GQ", _
		"GR", "HR", "HU", "IE", "IL", "IQ", "IR", "IS", "IT", _
		"JO", "KG", "KW", "KZ", "LB", "LI", "LR", "LT", "LU", _
		"LV", "MC", "MD", "MK", "MS", "MT", "NL", "NO", "OM", _
		"PL", "PS", "PT", "QA", "RO", "RU", "SA", "SE", "SI", _
		"SJ", "SK", "SL", "SM", "SO", "ST", "SY", "TD", "TJ", _
		"TM", "TR", "UA", "UZ", "VA", "YE", "YU", "UK", "EU"

			strTemp = "RIPE"
		
		' AFRINIC
		Case "AO", "BF", "BI", "BJ", "BW", "CD", "CF", "CI", "CM", "CY", _
		"DJ", "DZ", "EG", "ER", "ET", "GA", "GH", "GM", "GW", _
		"KE", "LS", "LY", "MA", "MG", "ML", "MR", "MU", "MW", "MZ", _
		"NA", "NA", "NE", "NG", "RW", "SC", "SD", "SN", "SZ", "TG", _
		"TN", "TZ", "UG", "ZA", "ZM", "ZW"

			strTemp = "AFRINIC"

		End Select
		
		ChooseIPWhois = strTemp

	End Function
	
	Private Function FormatIPAddress(strIPAddress, strRegistry, strClass)
		
		Dim strTemp
	
		If strRegistry <> "" Then
			strTemp = strTemp & "<a class=" & strClass & " " &_
				"href=""javascript:submitwhoisquery('" & strRegistry & "', '" & strIPAddress & "');"">"
		End If
		
		strTemp = strTemp & strIPAddress
						
		If strRegistry <> "" Then
			strTemp = strTemp & "</a>"
		End If
	
		FormatIPAddress = strTemp
	
	End Function
	
	Private Sub DisplayIPWhois()
	
		With Response
			.Write("<form name=""arin"" method=""get"" action=""http://ws.arin.net/whois/"" target=""_new"">")
			.Write("<input type=hidden name=""queryinput"" value=""""></form>" & vbcrlf)
			.Write("<form name=""apnic"" method=""post"" action=""http://wq.apnic.net/apnic-bin/whois.pl"" target=""_new"">")
			.Write("<input type=hidden name=form_type value=advanced>")
			.Write("<input type=hidden name=full_query_string>")
			.Write("<input type=hidden name=""searchtext""></form>" & vbcrlf)
			.Write("<form name=""ripe"" method=""get"" action=""http://www.ripe.net/perl/whois"" target=""_new"">")
			.Write("<input type=hidden name=form_type value=simple><input type=hidden name=""searchtext""></form>" & vbcrlf)
			.Write("<form name=""lacnic"" method=""post"" action=""http://lacnic.net/cgi-bin/lacnic/whois?lg=SP"" target=""_new"">")
			.Write("<input type=hidden name=""query""></form>" & vbcrlf)
			.Write("<form name=""afrinic"" method=""post"" action=""http://www.afrinic.net/cgi-bin/whois"" target=""_new"">")
			.Write("<input type=hidden name=""searchtext""></form>" & vbcrlf)
		End With

	End Sub
	
	Private Function DisplayReferrerLink(strData, strUrl, strClass)

		Dim strTemp
	
		If InStr(strData, "http://") > 0 Then
			strTemp = "<a target=_new class=" & strClass & " href=""" & Replace(strUrl, """", "%22") & """>" & strData & "</a>"
			
			If strData <> strUrl Then
				strTemp = strTemp & "&nbsp;<a href=""" & Replace(strUrl, """", "%22") & """ target=_new class=" & strClass & ">" &_
				"<img src=""images/link.gif"" alt=""" & Replace(strUrl, """", "%22") & """ border=0></a>"
			End If
			
		Else
			strTemp = "<a target=_new class=" & strClass & " href=""http://" & Replace(strUrl, """", "%22") & """>" & strData & "</a>"
			
			If strData <> strUrl Then
				strTemp = strTemp & "&nbsp;<a href=""" & Replace(strUrl, """", "%22") & """ target=_new class=" & strClass & ">" &_
				"<img src=""images/link.gif"" alt=""" & Replace(strUrl, """", "%22") & """ border=0></a>"
			End If
			
		End If
		
		DisplayReferrerLink = strTemp

	End Function
	
	Public Sub GeneratepresetdatesJS()
		
		Dim datStartDate, datEndDate
		
		With Response
			.Write("function presetdate()" & vbcrlf)
			.Write("{" & vbcrlf)

			datStartDate 	= Date()
			datEndDate 		= Date()
			.Write("if (document.report.preset.value == 'TODAY')" & vbcrlf)
			.Write("	{" & vbcrlf)
			.Write("	document.report.startdate.value='" & datStartDate & "';" & vbcrlf)
			.Write("	document.report.enddate.value='" & datEndDate & "';" & vbcrlf)
			.Write("	document.report.startdisplay.value='" & FormatDisplayDate(datStartDate, strShortDate) & "';" & vbcrlf)
			.Write("	document.report.enddisplay.value='" & FormatDisplayDate(datEndDate, strShortDate) & "';" & vbcrlf)
			.Write("	}" & vbcrlf)
			
			datStartDate 	= DateAdd("d", -1, Date())
			datEndDate 		= DateAdd("d", -1, Date())
			.Write("if (document.report.preset.value=='YESTERDAY')" & vbcrlf)
			.Write("	{" & vbcrlf)
			.Write("	document.report.startdate.value='" & datStartDate & "';" & vbcrlf)
			.Write("	document.report.enddate.value='" & datEndDate & "';" & vbcrlf)
			.Write("	document.report.startdisplay.value='" & FormatDisplayDate(datStartDate, strShortDate) & "';" & vbcrlf)
			.Write("	document.report.enddisplay.value='" & FormatDisplayDate(datEndDate, strShortDate) & "';" & vbcrlf)
			.Write("	}	" & vbcrlf)
			
			datStartDate 	= DateAdd("d", -7, Date())
			datEndDate 		= DateAdd("d", -1, Date())
			.Write("if (document.report.preset.value=='LASTWEEKROLL')" & vbcrlf)
			.Write("	{" & vbcrlf)
			.Write("	document.report.startdate.value='" & datStartDate & "';" & vbcrlf)
			.Write("	document.report.enddate.value='" & datEndDate & "';" & vbcrlf)
			.Write("	document.report.startdisplay.value='" & FormatDisplayDate(datStartDate, strShortDate) & "';" & vbcrlf)
			.Write("	document.report.enddisplay.value='" & FormatDisplayDate(datEndDate, strShortDate) & "';" & vbcrlf)
			.Write("	}	" & vbcrlf)
			
			datStartDate 	= DateSerial(Year(Date()), Month(Date()), 1)
			datEndDate 		= Date()
			.Write("if (document.report.preset.value=='THISMONTH')" & vbcrlf)
			.Write("	{" & vbcrlf)
			.Write("	document.report.startdate.value='" & datStartDate & "';" & vbcrlf)
			.Write("	document.report.enddate.value='" & datEndDate & "';" & vbcrlf)
			.Write("	document.report.startdisplay.value='" & FormatDisplayDate(datStartDate, strShortDate) & "';" & vbcrlf)
			.Write("	document.report.enddisplay.value='" & FormatDisplayDate(datEndDate, strShortDate) & "';" & vbcrlf)
			.Write("	}" & vbcrlf)

			datStartDate 	= DateSerial(Year(DateAdd("m", -1, Date())), Month(DateAdd("m", -1, Date())), 1)
			datEndDate 		= DateSerial(Year(DateAdd("m", -1, Date())), Month(DateAdd("m", -1, Date())), Day(DateAdd("d", 0 - Day(Date()), Date())))
			
			.Write("if (document.report.preset.value=='LASTMONTH')" & vbcrlf)
			.Write("	{" & vbcrlf)
			.Write("	document.report.startdate.value='" & datStartDate & "';" & vbcrlf)
			.Write("	document.report.enddate.value='" & datEndDate & "';" & vbcrlf)
			.Write("	document.report.startdisplay.value='" & FormatDisplayDate(datStartDate, strShortDate) & "';" & vbcrlf)
			.Write("	document.report.enddisplay.value='" & FormatDisplayDate(datEndDate, strShortDate) & "';" & vbcrlf)
			.Write("	}	" & vbcrlf)
			
			datStartDate 	= DateAdd("m", -1, Date())
			datEndDate 		= DateAdd("d", -1, Date())
			
			.Write("if (document.report.preset.value=='LASTMONTHROLL')" & vbcrlf)
			.Write("	{" & vbcrlf)
			.Write("	document.report.startdate.value='" & datStartDate & "';" & vbcrlf)
			.Write("	document.report.enddate.value='" & datEndDate & "';" & vbcrlf)
			.Write("	document.report.startdisplay.value='" & FormatDisplayDate(datStartDate, strShortDate) & "';" & vbcrlf)
			.Write("	document.report.enddisplay.value='" & FormatDisplayDate(datEndDate, strShortDate) & "';" & vbcrlf)
			.Write("	}				" & vbcrlf)
			.Write("}")
		End With

	End Sub
	
	Public Sub GenerateReportJS()
	
		With Response
			.Write("function GenerateReport(report)" & vbcrlf)
			.Write("{" & vbcrlf)
			.Write("var startdate = escape(document.report.startdate.value);" & vbcrlf)
			.Write("var enddate = escape(document.report.enddate.value);" & vbcrlf)
			.Write("var items = document.report.items.value;" & vbcrlf)
			.Write("var urlstr = 'default.asp?sd=' + startdate + '&ed=' + enddate + '&r=' + report + '&i=' + items;" & vbcrlf)
			.Write("document.location = urlstr;" & vbcrlf)
			.Write("}" & vbcrlf)
		End With
	
	End Sub
	
	Private Function GetReportArray()
	
		Dim aryTemp(27, 2)
		
		aryTemp(0,0) = "Summary"
		aryTemp(0,1) = "General"
		aryTemp(0,2) = 1
		aryTemp(1,0) = "Who's Online"
		aryTemp(1,1) = "General"
		aryTemp(1,2) = 1
		aryTemp(2,0) = "Pages"
		aryTemp(2,1) = "Pages & Files"
		aryTemp(2,2) = 1
		aryTemp(3,0) = "Urls"
		aryTemp(3,1) = "Pages & Files"
		aryTemp(3,2) = 1
		aryTemp(4,0) = "Daily"
		aryTemp(4,1) = "Pages & Files"
		aryTemp(4,2) = 1
		aryTemp(5,0) = "Hourly"
		aryTemp(5,1) = "Pages & Files"
		aryTemp(5,2) = 1
		aryTemp(6,0) = "By IP Address"
		aryTemp(6,1) = "Pages & Files"
		aryTemp(6,2) = 1
		aryTemp(7,0) = "Directories"
		aryTemp(7,1) = "Pages & Files"
		aryTemp(7,2) = 1
		aryTemp(8,0) = "File Types"
		aryTemp(8,1) = "Pages & Files"
		aryTemp(8,2) = 1
		aryTemp(9,0) = "Daily Unique Visitors"
		aryTemp(9,1) = "Visitors"
		aryTemp(9,2) = 1
		aryTemp(10,0) = "Daily Visits"
		aryTemp(10,1) = "Visitors"
		aryTemp(10,2) = 1
		aryTemp(11,0) = "Visits By IP Address"
		aryTemp(11,1) = "Visitors"
		aryTemp(11,2) = 1
		aryTemp(12,0) = "Entry Pages"
		aryTemp(12,1) = "Visitors"
		aryTemp(12,2) = 1
		aryTemp(13,0) = "Exit Pages"
		aryTemp(13,1) = "Visitors"
		aryTemp(13,2) = 1
		aryTemp(14,0) = "Domains"
		aryTemp(14,1) = "Referrers"
		aryTemp(14,2) = 1
		aryTemp(15,0) = "Hosts"
		aryTemp(15,1) = "Referrers"
		aryTemp(15,2) = 1
		aryTemp(16,0) = "Pages"
		aryTemp(16,1) = "Referrers"
		aryTemp(16,2) = 1
		aryTemp(17,0) = "URLs"
		aryTemp(17,1) = "Referrers"
		aryTemp(17,2) = 1
		aryTemp(18,0) = "Extensions"
		aryTemp(18,1) = "Referrers"
		aryTemp(18,2) = 1
		aryTemp(19,0) = "Searches"
		aryTemp(19,1) = "Search Engines"
		aryTemp(19,2) = 1
		aryTemp(20,0) = "Keywords"
		aryTemp(20,1) = "Search Engines"
		aryTemp(20,2) = 1
		aryTemp(21,0) = "Robots"
		aryTemp(21,1) = "Search Engines"
		aryTemp(21,2) = 0
		aryTemp(22,0) = "Country of Origin"
		aryTemp(22,1) = "Demographics"
		aryTemp(22,2) = 1
		aryTemp(23,0) = "Browsers"
		aryTemp(23,1) = "Demographics"
		aryTemp(23,2) = 1
		aryTemp(24,0) = "Operating Systems"
		aryTemp(24,1) = "Demographics"
		aryTemp(24,2) = 1
		aryTemp(25,0) = "Languages"
		aryTemp(25,1) = "Demographics"
		aryTemp(25,2) = 1
		aryTemp(26,0) = "Screen Area"
		aryTemp(26,1) = "Demographics"
		aryTemp(26,2) = 1
		aryTemp(27,0) = "User Agents"
		aryTemp(27,1) = "Demographics"
		aryTemp(27,2) = 1
		
		GetReportArray = aryTemp
		
	End Function
	
	Private Function ConvertCountryCode(strCode)
	
		Dim strTemp
	
		Select Case UCase(strCode)
		Case "AF"
			strTemp = "AFGHANISTAN"
		Case "AX"
			strTemp = "ÅLAND ISLANDS"
		Case "AL"
			strTemp = "ALBANIA"
		Case "DZ"
			strTemp = "ALGERIA"
		Case "AS"
			strTemp = "AMERICAN SAMOA"
		Case "AD"
			strTemp = "ANDORRA"
		Case "AO"
			strTemp = "ANGOLA"
		Case "AI"
			strTemp = "ANGUILLA"
		Case "AQ"
			strTemp = "ANTARCTICA"
		Case "AG"
			strTemp = "ANTIGUA AND BARBUDA"
		Case "AR"
			strTemp = "ARGENTINA"
		Case "AM"
			strTemp = "ARMENIA"
		Case "AP"
			strTemp = "African Regional Industrial Property Organization"
		Case "AW"
			strTemp = "ARUBA"
		Case "AU"
			strTemp = "AUSTRALIA"
		Case "AT"
			strTemp = "AUSTRIA"
		Case "AZ"
			strTemp = "AZERBAIJAN"
		Case "BS"
			strTemp = "BAHAMAS"
		Case "BH"
			strTemp = "BAHRAIN"
		Case "BD"
			strTemp = "BANGLADESH"
		Case "BB"
			strTemp = "BARBADOS"
		Case "BY"
			strTemp = "BELARUS"
		Case "BE"
			strTemp = "BELGIUM"
		Case "BZ"
			strTemp = "BELIZE"
		Case "BJ"
			strTemp = "BENIN"
		Case "BM"
			strTemp = "BERMUDA"
		Case "BT"
			strTemp = "BHUTAN"
		Case "BO"
			strTemp = "BOLIVIA"
		Case "BA"
			strTemp = "BOSNIA AND HERZEGOVINA"
		Case "BW"
			strTemp = "BOTSWANA"
		Case "BV"
			strTemp = "BOUVET ISLAND"
		Case "BR"
			strTemp = "BRAZIL"
		Case "IO"
			strTemp = "BRITISH INDIAN OCEAN TERRITORY"
		Case "BN"
			strTemp = "BRUNEI DARUSSALAM"
		Case "BG"
			strTemp = "BULGARIA"
		Case "BF"
			strTemp = "BURKINA FASO"
		Case "BI"
			strTemp = "BURUNDI"
		Case "KH"
			strTemp = "CAMBODIA"
		Case "CM"
			strTemp = "CAMEROON"
		Case "CA"
			strTemp = "CANADA"
		Case "CV"
			strTemp = "CAPE VERDE"
		Case "KY"
			strTemp = "CAYMAN ISLANDS"
		Case "CF"
			strTemp = "CENTRAL AFRICAN REPUBLIC"
		Case "TD"
			strTemp = "CHAD"
		Case "CL"
			strTemp = "CHILE"
		Case "CN"
			strTemp = "CHINA"
		Case "CX"
			strTemp = "CHRISTMAS ISLAND"
		Case "CC"
			strTemp = "COCOS (KEELING) ISLANDS"
		Case "CO"
			strTemp = "COLOMBIA"
		Case "KM"
			strTemp = "COMOROS"
		Case "CG"
			strTemp = "CONGO"
		Case "CD"
			strTemp = "CONGO, THE DEMOCRATIC REPUBLIC OF THE"
		Case "CK"
			strTemp = "COOK ISLANDS"
		Case "CR"
			strTemp = "COSTA RICA"
		Case "CI"
			strTemp = "COTE D'IVOIRE"
		Case "HR"
			strTemp = "CROATIA"
		Case "CU"
			strTemp = "CUBA"
		Case "CY"
			strTemp = "CYPRUS"
		Case "CZ"
			strTemp = "CZECH REPUBLIC"
		Case "DK"
			strTemp = "DENMARK"
		Case "DJ"
			strTemp = "DJIBOUTI"
		Case "DM"
			strTemp = "DOMINICA"
		Case "DO"
			strTemp = "DOMINICAN REPUBLIC"
		Case "EC"
			strTemp = "ECUADOR"
		Case "EG"
			strTemp = "EGYPT"
		Case "EU"
			strTemp = "EUROPEAN UNION"
		Case "SV"
			strTemp = "EL SALVADOR"
		Case "GQ"
			strTemp = "EQUATORIAL GUINEA"
		Case "ER"
			strTemp = "ERITREA"
		Case "EE"
			strTemp = "ESTONIA"
		Case "ET"
			strTemp = "ETHIOPIA"
		Case "FK"
			strTemp = "FALKLAND ISLANDS (MALVINAS)"
		Case "FO"
			strTemp = "FAROE ISLANDS"
		Case "FJ"
			strTemp = "FIJI"
		Case "FI"
			strTemp = "FINLAND"
		Case "FR"
			strTemp = "FRANCE"
		Case "GF"
			strTemp = "FRENCH GUIANA"
		Case "PF"
			strTemp = "FRENCH POLYNESIA"
		Case "TF"
			strTemp = "FRENCH SOUTHERN TERRITORIES"
		Case "GA"
			strTemp = "GABON"
		Case "GM"
			strTemp = "GAMBIA"
		Case "GE"
			strTemp = "GEORGIA"
		Case "DE"
			strTemp = "GERMANY"
		Case "GH"
			strTemp = "GHANA"
		Case "GI"
			strTemp = "GIBRALTAR"
		Case "GR"
			strTemp = "GREECE"
		Case "GL"
			strTemp = "GREENLAND"
		Case "GD"
			strTemp = "GRENADA"
		Case "GP"
			strTemp = "GUADELOUPE"
		Case "GU"
			strTemp = "GUAM"
		Case "GT"
			strTemp = "GUATEMALA"
		Case "GN"
			strTemp = "GUINEA"
		Case "GW"
			strTemp = "GUINEA-BISSAU"
		Case "GY"
			strTemp = "GUYANA"
		Case "HT"
			strTemp = "HAITI"
		Case "HM"
			strTemp = "HEARD ISLAND AND MCDONALD ISLANDS"
		Case "VA"
			strTemp = "HOLY SEE (VATICAN CITY STATE)"
		Case "HN"
			strTemp = "HONDURAS"
		Case "HK"
			strTemp = "HONG KONG"
		Case "HU"
			strTemp = "HUNGARY"
		Case "IS"
			strTemp = "ICELAND"
		Case "IN"
			strTemp = "INDIA"
		Case "ID"
			strTemp = "INDONESIA"
		Case "IR"
			strTemp = "IRAN, ISLAMIC REPUBLIC OF"
		Case "IQ"
			strTemp = "IRAQ"
		Case "IE"
			strTemp = "IRELAND"
		Case "IL"
			strTemp = "ISRAEL"
		Case "IT"
			strTemp = "ITALY"
		Case "JM"
			strTemp = "JAMAICA"
		Case "JP"
			strTemp = "JAPAN"
		Case "JO"
			strTemp = "JORDAN"
		Case "KZ"
			strTemp = "KAZAKHSTAN"
		Case "KE"
			strTemp = "KENYA"
		Case "KI"
			strTemp = "KIRIBATI"
		Case "KP"
			strTemp = "KOREA, DEMOCRATIC PEOPLE'S REPUBLIC OF"
		Case "KR"
			strTemp = "KOREA, REPUBLIC OF"
		Case "KW"
			strTemp = "KUWAIT"
		Case "KG"
			strTemp = "KYRGYZSTAN"
		Case "LA"
			strTemp = "LAO PEOPLE'S DEMOCRATIC REPUBLIC"
		Case "LV"
			strTemp = "LATVIA"
		Case "LB"
			strTemp = "LEBANON"
		Case "LS"
			strTemp = "LESOTHO"
		Case "LR"
			strTemp = "LIBERIA"
		Case "LY"
			strTemp = "LIBYAN ARAB JAMAHIRIYA"
		Case "LI"
			strTemp = "LIECHTENSTEIN"
		Case "LT"
			strTemp = "LITHUANIA"
		Case "LU"
			strTemp = "LUXEMBOURG"
		Case "MO"
			strTemp = "MACAO"
		Case "MK"
			strTemp = "MACEDONIA, THE FORMER YUGOSLAV REPUBLIC OF"
		Case "MG"
			strTemp = "MADAGASCAR"
		Case "MW"
			strTemp = "MALAWI"
		Case "MY"
			strTemp = "MALAYSIA"
		Case "MV"
			strTemp = "MALDIVES"
		Case "ML"
			strTemp = "MALI"
		Case "MT"
			strTemp = "MALTA"
		Case "MH"
			strTemp = "MARSHALL ISLANDS"
		Case "MQ"
			strTemp = "MARTINIQUE"
		Case "MR"
			strTemp = "MAURITANIA"
		Case "MU"
			strTemp = "MAURITIUS"
		Case "YT"
			strTemp = "MAYOTTE"
		Case "MX"
			strTemp = "MEXICO"
		Case "FM"
			strTemp = "MICRONESIA, FEDERATED STATES OF"
		Case "MD"
			strTemp = "MOLDOVA, REPUBLIC OF"
		Case "MC"
			strTemp = "MONACO"
		Case "MN"
			strTemp = "MONGOLIA"
		Case "MS"
			strTemp = "MONTSERRAT"
		Case "MA"
			strTemp = "MOROCCO"
		Case "MZ"
			strTemp = "MOZAMBIQUE"
		Case "MM"
			strTemp = "MYANMAR"
		Case "NA"
			strTemp = "NAMIBIA"
		Case "NR"
			strTemp = "NAURU"
		Case "NP"
			strTemp = "NEPAL"
		Case "NL"
			strTemp = "NETHERLANDS"
		Case "AN"
			strTemp = "NETHERLANDS ANTILLES"
		Case "NC"
			strTemp = "NEW CALEDONIA"
		Case "NZ"
			strTemp = "NEW ZEALAND"
		Case "NI"
			strTemp = "NICARAGUA"
		Case "NE"
			strTemp = "NIGER"
		Case "NG"
			strTemp = "NIGERIA"
		Case "NU"
			strTemp = "NIUE"
		Case "NF"
			strTemp = "NORFOLK ISLAND"
		Case "MP"
			strTemp = "NORTHERN MARIANA ISLANDS"
		Case "NO"
			strTemp = "NORWAY"
		Case "OM"
			strTemp = "OMAN"
		Case "PK"
			strTemp = "PAKISTAN"
		Case "PW"
			strTemp = "PALAU"
		Case "PS"
			strTemp = "PALESTINIAN TERRITORY, OCCUPIED"
		Case "PA"
			strTemp = "PANAMA"
		Case "PG"
			strTemp = "PAPUA NEW GUINEA"
		Case "PY"
			strTemp = "PARAGUAY"
		Case "PE"
			strTemp = "PERU"
		Case "PH"
			strTemp = "PHILIPPINES"
		Case "PN"
			strTemp = "PITCAIRN"
		Case "PL"
			strTemp = "POLAND"
		Case "PT"
			strTemp = "PORTUGAL"
		Case "PR"
			strTemp = "PUERTO RICO"
		Case "QA"
			strTemp = "QATAR"
		Case "RE"
			strTemp = "REUNION"
		Case "RO"
			strTemp = "ROMANIA"
		Case "RU"
			strTemp = "RUSSIAN FEDERATION"
		Case "RW"
			strTemp = "RWANDA"
		Case "SH"
			strTemp = "SAINT HELENA"
		Case "KN"
			strTemp = "SAINT KITTS AND NEVIS"
		Case "LC"
			strTemp = "SAINT LUCIA"
		Case "PM"
			strTemp = "SAINT PIERRE AND MIQUELON"
		Case "VC"
			strTemp = "SAINT VINCENT AND THE GRENADINES"
		Case "WS"
			strTemp = "SAMOA"
		Case "SM"
			strTemp = "SAN MARINO"
		Case "ST"
			strTemp = "SAO TOME AND PRINCIPE"
		Case "SA"
			strTemp = "SAUDI ARABIA"
		Case "SN"
			strTemp = "SENEGAL"
		Case "CS"
			strTemp = "SERBIA AND MONTENEGRO"
		Case "SC"
			strTemp = "SEYCHELLES"
		Case "SL"
			strTemp = "SIERRA LEONE"
		Case "SG"
			strTemp = "SINGAPORE"
		Case "SK"
			strTemp = "SLOVAKIA"
		Case "SI"
			strTemp = "SLOVENIA"
		Case "SB"
			strTemp = "SOLOMON ISLANDS"
		Case "SO"
			strTemp = "SOMALIA"
		Case "ZA"
			strTemp = "SOUTH AFRICA"
		Case "GS"
			strTemp = "SOUTH GEORGIA AND THE SOUTH SANDWICH ISLANDS"
		Case "ES"
			strTemp = "SPAIN"
		Case "LK"
			strTemp = "SRI LANKA"
		Case "SD"
			strTemp = "SUDAN"
		Case "SR"
			strTemp = "SURINAME"
		Case "SJ"
			strTemp = "SVALBARD AND JAN MAYEN"
		Case "SZ"
			strTemp = "SWAZILAND"
		Case "SE"
			strTemp = "SWEDEN"
		Case "CH"
			strTemp = "SWITZERLAND"
		Case "SY"
			strTemp = "SYRIAN ARAB REPUBLIC"
		Case "TW"
			strTemp = "TAIWAN, REPUBLIC OF CHINA"
		Case "TJ"
			strTemp = "TAJIKISTAN"
		Case "TZ"
			strTemp = "TANZANIA, UNITED REPUBLIC OF"
		Case "TH"
			strTemp = "THAILAND"
		Case "TL"
			strTemp = "TIMOR-LESTE"
		Case "TG"
			strTemp = "TOGO"
		Case "TK"
			strTemp = "TOKELAU"
		Case "TO"
			strTemp = "TONGA"
		Case "TT"
			strTemp = "TRINIDAD AND TOBAGO"
		Case "TN"
			strTemp = "TUNISIA"
		Case "TR"
			strTemp = "TURKEY"
		Case "TM"
			strTemp = "TURKMENISTAN"
		Case "TC"
			strTemp = "TURKS AND CAICOS ISLANDS"
		Case "TV"
			strTemp = "TUVALU"
		Case "UG"
			strTemp = "UGANDA"
		Case "UA"
			strTemp = "UKRAINE"
		Case "AE"
			strTemp = "UNITED ARAB EMIRATES"
		Case "GB","UK"
			strTemp = "UNITED KINGDOM"
		Case "US"
			strTemp = "UNITED STATES"
		Case "UM"
			strTemp = "UNITED STATES MINOR OUTLYING ISLANDS"
		Case "UY"
			strTemp = "URUGUAY"
		Case "UZ"
			strTemp = "UZBEKISTAN"
		Case "VU"
			strTemp = "VANUATU"
		Case "VE"
			strTemp = "VENEZUELA"
		Case "VN"
			strTemp = "VIET NAM"
		Case "VG"
			strTemp = "VIRGIN ISLANDS, BRITISH"
		Case "VI"
			strTemp = "VIRGIN ISLANDS, U.S."
		Case "WF"
			strTemp = "WALLIS AND FUTUNA"
		Case "EH"
			strTemp = "WESTERN SAHARA"
		Case "YE"
			strTemp = "YEMEN"
		Case "YU"
			strTemp = "YUGOSLAVIA"
		Case "ZM"
			strTemp = "ZAMBIA"
		Case "ZW"
			strTemp = "ZIMBABWE"
		Case "00"
			strTemp = "PRIVATE"
		Case Else
			strTemp = strCode
		End Select
		
		ConvertCountryCode = strTemp
	
	End Function
	
	Private Function ConvertToLanguageName(strLanguage)
	
		Dim strTemp, strLanguageAbbrev
	
		If InStr(strLanguage, ",") > 0 Then
			strLanguageAbbrev=Trim(Left(strLanguage, InStr(strLanguage, ",")-1))
		Else
			strLanguageAbbrev=Trim(strLanguage)
		End If
		
		If InStr(strLanguageAbbrev, ";") > 0 Then
			strLanguageAbbrev=Trim(Left(strLanguage, InStr(strLanguage, ";")-1))
		End If
		
		Select Case LCase(strLanguageAbbrev)
			Case "af"
				strTemp = "Afrikaans"
			Case "sq"
				strTemp = "Albanian"
			Case "ar"
				strTemp = "Arabic"
			Case "ar-sa"
				strTemp = "Arabic (Saudi Arabia)"
		    Case "ar-iq"
				strTemp = "Arabic (Iraq)"
		    Case "ar-eg"
				strTemp = "Arabic (Egypt)"
		    Case "ar-ly"
				strTemp = "Arabic (Libya)"
		    Case "ar-dz"
				strTemp = "Arabic (Algeria)"
		    Case "ar-ma"
				strTemp = "Arabic (Morocco)"
		    Case "ar-tn"
				strTemp = "Arabic (Tunisia)"
		    Case "ar-om"
				strTemp = "Arabic (Oman)"
		    Case "ar-ye"
				strTemp = "Arabic (Yemen)"
		    Case "ar-sy"
				strTemp = "Arabic (Syria)"
		    Case "ar-jo"
				strTemp = "Arabic (Jordan)"
		    Case "ar-lb"
		    	strTemp = "Arabic (Lebanon)"
		    Case "ar-kw"
				strTemp = "Arabic (Kuwait)"
		    Case "ar-ae"
				strTemp = "Arabic (U.A.E.)"
		    Case "ar-bh"
				strTemp = "Arabic (Bahrain)"
		    Case "ar-qa"
				strTemp = "Arabic (Qatar)"
		    Case "eu"
				strTemp = "Basque"
		    Case "bg"
				strTemp = "Bulgarian"
		    Case "be"
				strTemp = "Belarusian"
		    Case "ca"
				strTemp = "Catalan"
		    Case "zh-tw"
				strTemp = "Chinese (Taiwan)"
		    Case "zh-cn"
				strTemp = "Chinese (PRC)"
		    Case "zh-hk"
				strTemp = "Chinese (Hong Kong)"
		    Case "zh-sg"
				strTemp = "Chinese (Singapore)"
			Case "zh"
				strTemp = "Chinese"
		    Case "hr"
				strTemp = "Croatian"
		    Case "cs"
				strTemp = "Czech"
		    Case "da"
				strTemp = "Danish"
		    Case "nl"
				strTemp = "Dutch (Standard)"
		    Case "nl-be"
				strTemp = "Dutch (Belgian)"
		    Case "en"
				strTemp = "English"
		    Case "en-us"
				strTemp = "English (United States)"
		    Case "en-gb"
				strTemp = "English (British)"
		    Case "en-au"
				strTemp = "English (Australian)"
		    Case "en-ca"
				strTemp = "English (Canadian)"
		    Case "en-nz"
				strTemp = "English (New Zealand)"
		    Case "en-ie"
				strTemp = "English (Ireland)"
		    Case "en-za"
				strTemp = "English (South Africa)"
		    Case "en-jm"
				strTemp = "English (Jamaica)"
		    Case "en"
				strTemp = "English (Caribbean)"
		    Case "en-bz"
				strTemp = "English (Belize)"
		    Case "en-tt"
				strTemp = "English (Trinidad)"
		    Case "et"
				strTemp = "Estonian"
		    Case "fo"
				strTemp = "Faeroese"
		    Case "fa"
				strTemp = "Farsi"
		    Case "fi"
				strTemp = "Finnish"
		    Case "fr", "fr-fr"
				strTemp = "French (Standard)"
		    Case "fr-be"
				strTemp = "French (Belgian)"
		    Case "fr-ca"
				strTemp = "French (Canadian)"
		    Case "fr-ch"
				strTemp = "French (Swiss)"
		    Case "fr-lu"
				strTemp = "French (Luxembourg)"
		    Case "mk"
				strTemp = "FYRO Macedonian"
		    Case "gd"
				strTemp = "Gaelic (Scots)"
		    Case "gd-ie"
				strTemp = "Gaelic (Irish)"
		    Case "de"
				strTemp = "German (Standard)"
		    Case "de-ch"
				strTemp = "German (Swiss)"
		    Case "de-at"
				strTemp = "German (Austrian)"
		    Case "de-lu"
				strTemp = "German (Luxembourg)"
		    Case "de-li"
				strTemp = "German (Liechtenstein)"
		    Case "e", "el"
				strTemp = "Greek"
		    Case "he"
				strTemp = "Hebrew"
		    Case "hi"
				strTemp = "Hindi"
		    Case "hu"
				strTemp = "Hungarian"
		    Case "is"
				strTemp = "Icelandic"
		    Case "id"
				strTemp = "Indonesian"
		    Case "it"
				strTemp = "Italian (Standard)"
		    Case "it-ch"
				strTemp = "Italian (Swiss)"
			Case "it-it"
				strTemp = "Italian"
		    Case "ja"
				strTemp = "Japanese"
		    Case "ko"
				strTemp = "Korean"
		    Case "ko"
				strTemp = "Korean (Johab)"
		    Case "lv"
				strTemp = "Latvian"
		    Case "lt"
				strTemp = "Lithuanian"
		    Case "ms"
				strTemp = "Malaysian"
		    Case "mt"
				strTemp = "Maltese"
		    Case "no"
				strTemp = "Norwegian (Bokmal)"
		    Case "no"
				strTemp = "Norwegian (Nynorsk)"
		    Case "pl"
				strTemp = "Polish"
		    Case "pt-br"
				strTemp = "Portuguese (Brazil)"
		    Case "pt"
				strTemp = "Portuguese (Portugal)"
		    Case "rm"
				strTemp = "Rhaeto-Romanic"
		    Case "ro"
				strTemp = "Romanian"
		    Case "ro-mo"
				strTemp = "Romanian (Moldavia)"
		    Case "ru"
				strTemp = "Russian"
		    Case "ru-mo"
				strTemp = "Russian (Moldavia)"
		    Case "sz"
				strTemp = "Sami (Lappish)"
		    Case "sr"
				strTemp = "Serbian (Cyrillic)"
		    Case "sr"
				strTemp = "Serbian (Latin)"
		    Case "sk"
				strTemp = "Slovak"
		    Case "s", "sl"
				strTemp = "Slovenian"
		    Case "sb"
				strTemp = "Sorbian"
		    Case "es", "es-es"
				strTemp = "Spanish (Spain - Traditional Sort)"
		    Case "es-mx"
				strTemp = "Spanish (Mexican)"
		    Case "es-gt"
				strTemp = "Spanish (Guatemala)"
		    Case "es-cr"
				strTemp = "Spanish (Costa Rica)"
		    Case "es-pa"
				strTemp = "Spanish (Panama)"
		    Case "es-do"
				strTemp = "Spanish (Dominican Republic)"
		    Case "es-ve"
				strTemp = "Spanish (Venezuela)"
		    Case "es-co"
				strTemp = "Spanish (Colombia)"
		    Case "es-pe"
				strTemp = "Spanish (Peru)"
		    Case "es-ar"
				strTemp = "Spanish (Argentina)"
		    Case "es-ec"
				strTemp = "Spanish (Ecuador)"
		    Case "es-cl"
				strTemp = "Spanish (Chile)"
		    Case "es-uy"
				strTemp = "Spanish (Uruguay)"
		    Case "es-py"
				strTemp = "Spanish (Paraguay)"
		    Case "es-bo"
				strTemp = "Spanish (Bolivia)"
		    Case "es-sv"
				strTemp = "Spanish (El Salvador)"
		    Case "es-hn"
				strTemp = "Spanish (Honduras)"
		    Case "es-ni"
				strTemp = "Spanish (Nicaragua)"
		    Case "es-pr"
				strTemp = "Spanish (Puerto Rico)"
		    Case "sx"
				strTemp = "Sutu"
		    Case "sv"
				strTemp = "Swedish"
		    Case "sv-fi"
				strTemp = "Swedish (Finland)"
		    Case "th"
				strTemp = "Thai"
		    Case "ts"
				strTemp = "Tsonga"
		    Case "tn"
				strTemp = "Tswana"
		    Case "tr"
				strTemp = "Turkish"
		    Case "uk"
				strTemp = "Ukrainian"
		    Case "ur"
				strTemp = "Urdu"
		    Case "ve"
				strTemp = "Venda"
		    Case "vi"
				strTemp = "Vietnamese"
			Case "xh"
				strTemp = "Xhosa"
			Case "ji"
				strTemp = "Yiddish"
			Case "zu"
				strTemp = "Zulu"
			Case Else
				strTemp = strLanguageAbbrev 
		End Select
		
		ConvertToLanguageName = strTemp
		
	End Function
	
End Class

%>