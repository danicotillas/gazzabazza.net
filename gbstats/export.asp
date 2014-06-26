<% Option Explicit 
%>
<!--#Include File="config.asp"-->
<!--#Include File="conn.asp"-->
<!--#Include File="core.asp"-->
<!--#Include File="clsReport.asp"-->
<%
Server.ScriptTimeout = 900

Dim strType : strType = UCase(Request.Querystring("type"))
Dim strReport : strReport = Request.Form("report")
Dim strSite : strSite = Request.Form("site")
Dim strDesc : strDesc = Request.Form("desc")
Dim intCols : intCols = Request.Form("cols")
Dim aryData1 : aryData1 = Split(Request.Form("data1"), ", ")
Dim aryData2 : aryData2 = Split(Request.Form("data2"), ", ")
Dim aryData3 : aryData3 = Split(Request.Form("data3"), ", ")
Dim aryData4 : aryData4 = Split(Request.Form("data4"), ", ")
Dim intTotal : intTotal = Request.Form("total")
Dim datExport : datExport = Request.Form("export")

Dim objExport : Set objExport = New MTReport
With ObjExport
	.Database			= aryMTDB
	.Config				= aryMTConfig
End With

' AUTHENTICATION
Call CreateDatabaseConnection(1)
Dim blnAdmin : blnAdmin = Authenticate(False, aryMTDB(5))
Call CloseDatabaseConnection()

Dim intLoop

Select Case strType

Case "CSV"
	With Response
		.ContentType="application/csv"
		.AddHeader "Content-Disposition", "filename=export_" & Replace(strReport, " ", "_") & ".csv"
		.Write(FormatExportData(strSite) & vbcrlf)
		.Write(FormatExportData(strReport) & vbcrlf)
		.Write(FormatExportData(strDesc) & vbcrlf & vbcrlf)
		For intLoop = 0 To UBound(aryData1)
			.Write(FormatExportData(aryData1(intLoop)) & ",")
			.Write(FormatExportData(aryData2(intLoop)) & ",")
			If intCols = 4 Then
				.Write(FormatExportData(aryData3(intLoop)) & ",")
				.Write(FormatExportData(aryData4(intLoop)) & vbcrlf)
			ElseIf intCols = 3 Then
				.Write(FormatExportData(aryData3(intLoop)) & vbcrlf)
			Else
				.Write(vbcrlf)
			End If
		Next
		
		If intTotal <> "" Then
			.Write("Total: ," & intTotal & vbcrlf)
		End If
		
		.Write(vbcrlf & "Report Generated at " & datExport & " by FJstats")
		
	End With
	
End Select

Function FormatExportData(strData)
	
	strData = Replace(strData, "%2C", ",")
	strData = Replace(strData, "%22", """")
		
	FormatExportData = """" & strData & """"
	
End Function
%>
