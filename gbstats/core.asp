<%
Function FormatDatabaseDate(datDate)

	Dim datDateTemp, datTimeTemp, strDateFormat, strTimeFormat
	Dim datTemp, strSeparator, datDatabaseDate, datDatabaseTime, datFull
	
	' SET DATABASE DATE AND TIME FORMATS
	If aryMTDB(0) = "MSSQL" Then
		strDateFormat = "YYYYMMDD"
	Else
		strDateFormat = "YYYY-MM-DD"
	End If
	
	strTimeFormat = "HH:MM:SS"
	
	' MAKE SURE FORMAT IS ALL UPPERCASE
	datDateTemp = UCase(strDateFormat)
	datTimeTemp = UCase(strTimeFormat)

	' BEGIN REPLACING TOKENS ON DATE
	datDateTemp = Replace(datDateTemp, "DD", FormatDatePart(Day(datDate)))
	datDateTemp = Replace(datDateTemp, "MMMM", MonthName(Month(datDate), False))
	datDateTemp = Replace(datDateTemp, "MMM", MonthName(Month(datDate), True))
	datDateTemp = Replace(datDateTemp, "MM", FormatDatePart(Month(datDate)))
	datDateTemp = Replace(datDateTemp, "YYYY", Year(datDate))
	datDateTemp = Replace(datDateTemp, "YY", Right(Year(datDate), 2))
	
	' BEGIN REPLACING TOKENS ON TIME
	datTimeTemp = Replace(datTimeTemp, "HH", FormatDatePart(DatePart("h", datDate)))
	datTimeTemp = Replace(datTimeTemp, "MM", FormatDatePart(DatePart("n", datDate)))
	datTimeTemp = Replace(datTimeTemp, "SS", FormatDatePart(DatePart("s", datDate)))
	
	If aryMTDB(0) = "MSACCESS" Then
		strSeparator = "#"
	Else
		strSeparator = "'"
	End If

	' BUILD FINAL DATE FORMAT
	datTemp = strSeparator & datDateTemp & " " & datTimeTemp & strSeparator

	FormatDatabaseDate = datTemp

End Function

Function FormatDatePart(datPart)
	Dim datTemp
	
		If Len(datPart) = 1 Then
			datTemp = "0" & datPart
		Else
			datTemp = datPart
		End If

	FormatDatePart = datTemp
End Function

Function FormatDatabaseString(strString, intLength)

	Dim strTemp
	
	strTemp = "'" & Replace(Left(strString, intLength), "'", "''") & "'"

	FormatDatabaseString = strTemp

End Function

Function Authenticate(blnRequireAdmin, strTablePrefix)
	
	Dim blnAdmin, intAuth : intAuth = 0
	Dim strUsername : strUsername = Request.Cookies("FJstats")("username")
	Dim strPassword : strPassword = Request.Cookies("FJstats")("password")

	If strUsername <> "" Then
	
		Dim strSql : strSql = "SELECT u_admin " &_
			"FROM " & strTablePrefix & "Users " &_
			"WHERE u_username = " & FormatDatabaseString(strUsername, 20) & " " &_
			"AND u_password = " & FormatDatabaseString(strPassword, 20)
	
		Dim rsAuth : Set rsAuth = Server.CreateObject("ADODB.Recordset")
		
		On Error Resume Next
	
		rsAuth.Open strSql, objConn, 1, 2, &H0001
		
		If Err.Number <> 0 Then
			Call DisplayDBConnError(Err)
		End If
		
		On Error Goto 0
		
		If Not rsAuth.Eof Then
			blnAdmin = CBool(rsAuth(0))
			If blnRequireAdmin = True Then
				If blnAdmin = True Then
					intAuth = 1
				Else
					intAuth = -1
				End If
			Else
				intAuth = 1
			End If
		End If
		rsAuth.Close : Set rsAuth = Nothing
	Else
		intAuth = -2
	End If
	
	If intAuth <> 1 Then
		Response.Redirect "login.asp?action=failure&code=" & intAuth
	End If
	
	Authenticate = blnAdmin
	
End Function

Sub CreateDatabaseConnection(intError)

	Dim strSql, strConn, strLocationType, strTemp, intPort, aryServer
	Dim blnPort : blnPort = False
	
	If InStr(aryMTDB(1), ":") > 0 And aryMTDB(0) <> "MSACCESS" Then
		aryServer = Split(aryMTDB(1), ":")
		aryMTDB(1) = aryServer(0)
		intPort = Int(aryServer(1))
		If intPort > 0 Then
			blnPort = True
		End If
	End If
	
	If aryMTDB(0) = "MSSQL" Then

		strConn = "DRIVER={SQL Server};" &_
			"SERVER=" & aryMTDB(1) & ";"
			If blnPort = True Then
				strConn = strConn & "PORT=" & intPort & ";"
			End If
			strConn = strConn & "DATABASE=" & aryMTDB(2) & ";" &_
			"UID=" & aryMTDB(3) & ";" &_
			"PWD=" & aryMTDB(4) & ";" &_
			"Provider=MSDASQL.1"
			
	ElseIf aryMTDB(0) = "MYSQL" Then

		strConn = "DRIVER={MySQL ODBC 3.51 Driver};" &_
			"SERVER=" & aryMTDB(1) & ";"
			If blnPort = True Then
				strConn = strConn & "PORT=" & intPort & ";"
			Else
				strConn = strConn & "PORT=3306;"
			End If
			strConn = strConn & "DATABASE=" & aryMTDB(2) & ";" &_
			"UID=" & aryMTDB(3) & ";" &_
			"PWD=" & aryMTDB(4) & ";Option=16387"

	Else
		If Len(aryMTDB(1)) > 2 Then
		If Mid(aryMTDB(1), 2, 1) = ":" Or Mid(aryMTDB(1), 1, 2) = "\\" Then
				strLocationType = "ABSOLUTE"
			Else
				strLocationType = "VIRTUAL"
			End If
		Else
			strLocationType = "VIRTUAL"
		End If
		
		If strLocationType = "ABSOLUTE" Then
			strTemp = aryMTDB(1) & "\" & aryMTDB(2)
		Else
			strTemp = Server.MapPath(aryMTDB(1) & "/" & aryMTDB(2))
		End If
		
		strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strTemp
	End If
	
	Set objConn = Server.CreateObject("ADODB.Connection")
	
	If intError = 0 Then
	
		objConn.Open strConn
		
	ElseIf intError = 1 Then
	
		On Error Resume Next
	
		objConn.Open strConn
		
		If Err.Number <> 0 Then
			Call DisplayDBConnError(Err)
		End If
		
		On Error Goto 0
		
	Else
		
		On Error Resume Next
		objConn.Open strConn
		On Error Goto 0
		
	End If

End Sub

Sub CloseDatabaseConnection()

	If IsObject(objConn) Then
		objConn.Close : Set objConn = Nothing
	End If

End Sub

Function ShowProductInfo()
	
	Dim strTemp
	
	strTemp = ""
	
	ShowProductInfo = strTemp
	
End Function

Function FormatDisplayDate(datDate, strFormat)
	
	Dim datTemp : datTemp = UCase(strFormat)
	
	datTemp = Replace(datTemp, "DDDD", WeekdayName(Weekday(datDate), False))
	datTemp = Replace(datTemp, "DDD", WeekdayName(Weekday(datDate), True))
	datTemp = Replace(datTemp, "DD", Day(datDate))
	
	datTemp = Replace(datTemp, "MMMM", MonthName(Month(datDate), False))
	datTemp = Replace(datTemp, "MMM", MonthName(Month(datDate), True))
	datTemp = Replace(datTemp, "MM", Month(datDate))
	
	datTemp = Replace(datTemp, "YYYY", Year(datDate))
	datTemp = Replace(datTemp, "YY", Right(Year(datDate), 2))

	FormatDisplayDate = datTemp

End Function

Sub DisplayDBConnError(Err)

	With Response
		.Write("<html><head><link rel=""stylesheet"" href=""style.css"" type=""text/css""></head>")
		.Write("<body style=""padding: 10px;"">")
		.Write("<table border=0 cellpadding=5 cellspacing=0 class=settings><tr><td>")
		.Write("<p class=error>There was an error connecting to the database: </p>")
		.Write("<p>Number: " & Err.Number & "<br>")
		.Write("Source: " & Err.Source & "<br>")
		.Write("Description: " & Err.Description & "</p>")
		.Write("<p>If you have not setup FJstats yet, go to the ")
		.Write("<a href=""setup.asp"">setup</a> page.</p>")
		.Write("</td></tr></table>")
		.Write("</body></html>")
	End With
	
	Response.End

End Sub
%>