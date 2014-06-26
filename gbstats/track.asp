<% 'Option Explicit %>
<!--#Include File="config.asp"-->
<!--#Include File="conn.asp"-->
<!--#Include File="core.asp"-->
<!--#Include File="clsLog.asp"-->
<%

' GET SOME QUERYSTRING VARS
Dim strUrl : strUrl	= Request.Querystring("mtr")
Dim blnImage : blnImage = CBool(Request.Querystring("mti"))

' CHECK WHAT TYPE OF LOGGING METHOD IS BEING USED
' 0 - ASP EXECUTE METHOD
' 1 - REDIRECT FILE METHOD
' 2 - JAVASCRIPT METHOD

' SET LOGGING TYPE IN CASE UNSPECIFIED
Dim intType
If strUrl <> "" Then
	intType		= 1
Else
	intType		= 0
End If

' GET LOGGING TYPE IF SPECIFIED
If Request.Querystring("mtt") <> "" Then 
	intType = Request.Querystring("mtt")
End If

' GET SCREENAREA IF AVAILABLE
Dim strScreenarea
If Request.Querystring("mts") <> "x" Then
	strScreenArea = Request.Querystring("mts")
End If

Dim blnExclude
If Request.Cookies("tosh_exclude") <> "" Then
	blnExclude = True
Else
	blnExclude = False
End If

' GET PAGE NAME
Dim strPageTitle
If intType = 0 Then
	strPageTitle = Request.Cookies("Tosh")("pagetitle")
ElseIf intType = 2 Then
	strPageTitle = Request.Querystring("mtpt")
End If

' LOG REQUEST IF LOGGING IS ENABLED
If (aryMTConfig(2) = True Or aryMTConfig(2) = "") And blnExclude = False Then

	' INSTANTIATE OBJECT FROM CLASS.ASP FILE
	Dim objTrack : Set objTrack = New MTLog
	
	Call CreateDatabaseConnection(0)
	
	' SET SOME PROPERTIES
	With ObjTrack
		.Database		= aryMTDB
		.Config			= aryMTConfig
		.PageTitle		= strPageTitle
	End With

	' CHECK TO SEE IF IP MATCHES LOG EXCLUSION LIST
	If Not objTrack.MatchIPAddress(aryMTConfig(3)) Then
		' PERFORM LOGGING OPERATION
		Call objTrack.LogFile(strUrl, intType, strScreenArea)
	End If
	
	Set objTrack = Nothing
	
	Call CloseDatabaseConnection()

End If

' REDIRECT TO PAGE IF USING REDIRECT FILE METHOD (intType = 1)
If CInt(intType) = 1 Then
	Response.Redirect strUrl
End If

If blnImage = True Then
	Response.ContentType ="image/gif"
%>
<!--#Include File="images/spacer.gif"-->
<% End If %>
