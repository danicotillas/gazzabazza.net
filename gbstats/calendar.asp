<% Option Explicit
%>
<!--#Include File="config.asp"-->
<!--#Include File="core.asp"-->
<%
Dim strClass, intDayLoop

Dim strShortDate : strShortDate = aryMTConfig(10)

Dim strFormName : strFormname = Request.Querystring("name")
Dim datSelected : datSelected = Request.Querystring("sdate")
Dim datCurrent : datCurrent	= Request.Querystring("cdate")

If Isdate(datSelected) = False Then
	datSelected = Date()
End If

If Isdate(datCurrent) = False Then
	datCurrent = datSelected
End If

datSelected = Cdate(datSelected)
datCurrent = Cdate(datCurrent)

Dim datToday : datToday	= Date()
Dim datMonth : datMonth	= Month(datCurrent)
Dim datMonthName : datMonthName	= Monthname(datMonth)
Dim datYear : datYear = Year(datCurrent)
Dim datNextMonth : datNextMonth	= Month(DateAdd("m", 1, datCurrent))
Dim datNextMonthYear : datNextMonthYear	= Year(DateAdd("m", 1, datCurrent))
Dim intDaysInMonth : intDaysInMonth = Day(DateSerial(Year(datCurrent), Month(datCurrent) + 1, 0))
Dim intStartWeekday : intStartWeekday = Weekday(DateAdd("d", (1 - Day(datCurrent)), datCurrent), 1)
Dim intEndWeekday : intEndWeekday = Weekday(DateAdd("d", intDaysInMonth, DateAdd("d", (0 - Day(datCurrent)), datCurrent)), 1)
Dim intWeekdayPos : intWeekdayPos = 0

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">

<html>
<head>
	<title>Calendar: Select a date</title>
	<style>
	body 			{font-family:verdana, helvetica; font-size: 12px; margin: 2px;}
	td				{font-size: 12px;}
	td.header 		{font-weight: bold; text-align: center; color: #FFFFFF; background-color: #203759;}
	td.today 		{color: #FFFFFF; background-color: #CCCCCC; }
	td.normal 		{background-color: #EEEEEE; }
	td.blank 		{background-color: #FFFFFF; }
	td.selected 	{background-color: #CCCCCC; }
	th 				{font-weight: bold; background-color: #E6E9ED; font-size: 12px; }
	a:link 			{color: #000000; text-decoration: none} 
	a:visited 		{color: #000000; text-decoration: none} 
	a:hover 		{color: #FF6633; text-decoration: none;} 
	a:active 		{color: #000000; text-decoration: none}
	a.month:link 	{color: #EEEEEE; text-decoration: none} 
	a.month:visited {color: #EEEEEE; text-decoration: none}
	a.month:hover 	{color: #FF6633; text-decoration: none} 
	a.month:active 	{color: #000000; text-decoration: none}
	</style>
	<script language="JavaScript">
		function selectdate(datNormal, datFormat) 
		{
		opener.document.<%=strFormname%>date.value=datNormal;
		opener.document.<%=strFormname%>display.value=datFormat;
		<% If InStr(strFormname, "report") Then %>
		opener.document.report.preset.value='CUSTOM';
		<% End If %>
		self.close();
		}
	</script>
</head>
<body>
<%
With Response
	.Write "<table border=1 cellpadding=2 cellspacing=0>"
	.Write "<tr><td colspan=7 class=header><table border=0 cellpadding=0 cellspacing=0 width=""100%"">"
	.Write "<form method=post>"
	.Write "<tr><td class=header align=left>"
	.Write "<a class=month href=""?cdate=" & Dateadd("m", -1, datCurrent) & "&sdate=" & datSelected & "&name=" & strFormname & """>"
	.Write "&nbsp;<strong>&laquo;</strong>&nbsp;</a>"
	.Write "</td><td class=header align=center>" & datMonthName & " " & datYear & "</td>"
	.Write "<td class=header align=right>"
	.Write "<a class=month href=""?cdate=" & Dateadd("m", 1, datCurrent) & "&sdate=" & datSelected & "&name=" & strFormname & """>"
	.Write "&nbsp;<strong>&raquo;</strong>&nbsp;</a>"
	.Write "</td></tr>"
	.Write "</form></table></tr>"
	.Write "<tr align=center><th>S</th><th>M</th><th>T</th><th>W</th><th>T</th><th>F</th><th>S</th></tr>"
	.Write "<tr align=right>"
End With

If intStartWeekday > 1 Then
	For intDayLoop = 1 to (intStartWeekday - 1)
		intWeekdayPos = intWeekdayPos + 1
		Response.Write "<td class=blank>&nbsp;</td>"
	Next
End If

For intDayLoop = 1 to intDaysInMonth
	intWeekdayPos = intWeekdayPos + 1
	
	' START NEW ROW IF BEGINNING OF WEEK
	If intWeekdayPos Mod 7 = 1 Then 
		Response.Write "<tr align=right>"
	End If
	
	if datSelected = DateSerial(datYear, datMonth, intDayLoop) Then
		strClass = "selected"
	elseif datToday = DateSerial(datYear, datMonth, intDayLoop) Then
		strClass = "today"
	Else
		strClass = "normal"
	End If
	
	Response.Write "<td class=" & strClass & ">"
	Response.Write "<a href=""javascript: selectdate('" & DateSerial(datYear, datMonth, intDayLoop) & "',"
	Response.Write "'" & FormatDisplayDate(DateSerial(datYear, datMonth, intDayLoop), strShortDate) & "');"">"
	Response.Write intDayLoop & "</a></td>"
	
	' CLOSE OFF ROW IF END OF WEEK
	If intWeekdayPos Mod 7 = 0 Then 
		Response.Write "</tr>"
	End If
Next

If intEndWeekday < 7 Then
	For intDayLoop = 1 to (7 - intEndWeekday)
		intWeekdayPos = intWeekdayPos + 1
		Response.Write "<td class=blank>&nbsp;</td>"
	Next
End If

Response.Write "</table>"
Response.Write Request.Form("cdate")

Function FormatDate (intDay, intMonth, intYear)

	Dim datTemp, datReference
	
	datReference = Date()
	datTemp = datReference
	
	' SET YEAR
	datTemp = DateAdd("yyyy", intYear - Year(datReference), datTemp)
	' SET MONTH
	datTemp = DateAdd("m", intMonth - Month(datReference), datTemp)
	' SET DAY
	datTemp = DateAdd("d", intDay - Day(datReference), datTemp)

	FormatDate = datTemp

End Function

%>
</body>
</html>
