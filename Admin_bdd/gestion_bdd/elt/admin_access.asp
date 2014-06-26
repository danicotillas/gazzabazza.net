<%
' ±AxS 1.09 - WEB MANAGEMENT FOR MSACCESS DATABASES - www.ASP-PHP.net
' by Didier YVER - webmaster@asp-php.net (please report any bug !)
' http://www.asp-php.net/scripts/asp-php/axs.php

Version = "±AxS 1.09 © ASP-PHP.net 2003"
Set FSO = Server.CreateObject("Scripting.FileSystemObject")
URL = lcase(request.serverVariables("SCRIPT_NAME"))

NAV = "IE" : if inStr(Request.ServerVariables("HTTP_USER_AGENT"),"MSIE")=0 then NAV="NN"
if inStr(Request.ServerVariables("HTTP_USER_AGENT"),"Netscape6") then NAV="N6"
SIZEINPUT = 20 : if NAV <> "IE" then SIZEINPUT = 15

' --- USER PARAMS / PARAMETRES UTILISATEUR ---
%><!-- #include file="../_params.asp"--><%

if user_debug then
	server.scriptTimeout=1
else
	on error resume next
end if

' --- PROTECTION ---
if request.form("login")=user_login and request.form("pass")=user_pass then
	session("AxS_login")=request.form("login")
	session("AxS_pass")=request.form("pass")
end if	
user_ok = ((session("AxS_login")=user_login and session("AxS_pass")=user_pass) or user_login="")

' --- READ DICTIONARY / LECTURE DU DICTIONNAIRE ---
%><!-- #include file="../lng/_read.asp"--><%

delimg = "<img src=""minipoub.gif"" width=14 height=15 border=0 title=""" & trad(30) & """>"

' --- DATABASES MANAGEMENT  / GESTION DES BD ---
%><!-- #include file="../bdd/_bdd.asp"--><%

SQLpath = user_baseFolder & baseRep & BASE & ".sql"

TABLE = request.form("table") : if TABLE="" then TABLE = request("table")

' --- FIELDS TYPE / TYPE DES CHAMPS ---

function FieldType(champ,lng)
	select case champ
		case  -1 : if lng then FieldType=trad(36) else FieldType="Counter"
		case   2 : if lng then FieldType=trad(16) else FieldType="SmallInt"
		case   3 : if lng then FieldType=trad(17) else FieldType="Integer"
		case   4 : if lng then FieldType=trad(18) else FieldType="Real"
		case   5 : if lng then FieldType=trad(19) else FieldType="Number"
		case   6 : if lng then FieldType=trad(20) else FieldType="Money"
		case  11 : if lng then FieldType=trad(21) else FieldType="YesNo"
		case  17 : if lng then FieldType=trad(22) else FieldType="Byte"
		case 128 : if lng then FieldType=trad(24) else FieldType="Binary"
		case 131 : if lng then FieldType=trad(25) else FieldType="???" ' Numeric
		case 135 : if lng then FieldType=trad(26) else FieldType="Date"
		case 200,202 : if lng then FieldType=trad(27) else FieldType="Text"
		case 201,203 : if lng then FieldType=trad(28) else FieldType="Memo"
		case 205 : if lng then FieldType=trad(29) else FieldType="LongBinary"
		case else : FieldType=cStr(champ)
	end select
end function

' --- SQL MANAGEMENT  / GESTION DES REQUETES ---
SQLstr0 = request.form("AxS_strsql")
SQLstr = replace(SQLstr0,VbCrLf," ")
%><!-- #include file="../mod/edit/sql_top.asp"--><%

	' --- NEW TABLE / NOUVELLE TABLE ---
	if request.form("newtable")<>"" then
		Conn.execute("CREATE TABLE [" & request.form("newtable") & "] (ID Counter)")
		TABLE = request.form("newtable")

	elseif TABLE<>"" then

		' --- EXPORT TABLE ---
		if request("export")<>"" then
			%><!-- #include file="../mod/edit/export/export.asp"--><%
		' --- NEW FIELD / NOUVEAU CHAMP ---
		elseif request.form("newChamp")<>"" then
			SQL = "ALTER TABLE [" & TABLE & "] ADD [" & request.form("newChamp") & "] " _
				& request("typeChamp")
			if request("typeChamp")="Text" and request("sizeChamp")<>"" _
			then SQL = SQL & "(" & request("sizeChamp") & ")"
			Conn.execute(SQL)
			' --- RENAME FIELD - RENOMME CHAMP ---
			if request("renamefld")<>"" and err=0 then
				SQL = "UPDATE [" & TABLE & "] SET [" & request.form("newChamp") & "]=["& request("renamefld") & "]"
				Conn.execute(SQL)
				SQL = "ALTER TABLE [" & TABLE & "] DROP COLUMN [" & request("renamefld") & "]"
				Conn.execute(SQL)
			end if

		' --- NEW RECORD / NOUVELLE FICHE ---
		elseif request.queryString("newfi")<>"" then
			SQL = "SELECT * FROM [" & TABLE & "]"
			RS.open SQL,conn,1,2
			RS.addNew
			RS.update
			SQLstr = "SELECT * FROM [" & TABLE & "] WHERE [" _
				& request.queryString("newfi") & "]=" & RS(cstr(request.queryString("newfi")))
			RS.close

		' --- MODIFY/DELETE RECORD / MODIFIE/DETRUIT LA FICHE ---
		elseif request.form("idname")<>"" then
			if request.form("submit") = trad(49) then
				SQL = "DELETE * FROM [" & TABLE & "] WHERE [" & request.form("idname") & "]=" & request.form("idval")
				conn.execute(SQL)
			else
				SQL = "SELECT * FROM [" & TABLE & "] WHERE [" & request.form("idname") & "]=" & request.form("idval")
				RS.open SQL,conn,3,3
				for each Champ in RS.fields
					if not isempty(request.form(Champ.name)) then
						select case Champ.type
							case 2,3,4,5,6,17,135:
								if request.form(Champ.name)<>"" then Champ.value=request.form(Champ.name)
							case 11: Champ.value=(request.form(Champ.name)="True")
							case else Champ.value=request.form(Champ.name)
						end select
					end if
				next
				RS.update : RS.close
			end if

		' --- DELETE RECORDS / DETRUIT LES FICHES ---
		elseif request.form("boxdel")<>"" and request.queryString("record")="" then
			IDfld = request.form("IDfld")
			SQL = "DELETE * FROM [" & TABLE & "] WHERE [" & IDfld & "]=0"
			for i = 1 to request.form("boxdel").count
				SQL = SQL & " OR [" & IDfld & "]=" & request.form("boxdel").item(i)
			next
			conn.execute(SQL)

		' --- SEARCH / RECHERCHE ---
		elseif request.form("fldSrch")<>"" then
			SQLstr = "SELECT * FROM [" & TABLE & "] WHERE [" & request.form("fldSrch") & "] "
			tempSrch = replace(request.form("search"),"'","''")
			if left(tempSrch,1)="=" then
				SQLstr = SQLstr & "='" & mid(tempSrch,2) & "'"
			elseif left(tempSrch,1)=">" then
				SQLstr = SQLstr & "LIKE '%" & mid(tempSrch,2) & "%'"
			else
				tempSrch=trim(request.form("Search"))
				while instr(tempSrch,"  ")
					tempSrch = replace(tempSrch,"  "," ")
				wend
				tempSrch = replace(tempSrch," ","%' OR [" & request.form("fldSrch") & "] LIKE '%")
				tempSrch = replace(tempSrch,"+","%' AND [" & request.form("fldSrch") & "] LIKE '%")
				tempSrch = replace(tempSrch,"-","%' AND [" & request.form("fldSrch") & "] NOT LIKE '%")
				SQLstr = SQLstr & "LIKE '%" & tempSrch & "%'"
			end if
			if request.form("idSrch")<>"" then
				SQLstr = SQLstr & " ORDER BY [" & request.form("idSrch") & "] DESC"
			end if
		end if

	end if

	' --- VERIFICATION  ---
	Set ObjTable = Conn.OpenSchema(20)
	do while not ObjTable.eof
		if ObjTable("table_type")="TABLE" then
			if Table1="" then Table1=ObjTable("table_name")
			if TABLE = ObjTable("table_name") then exit do
		end if
		ObjTable.moveNext
	loop
	if ObjTable.eof then TABLE = Table1
	ObjTable.close
end if

' --- CORRECT URL / CORRECTION URL ---
urlk = replace("&base=" & baseRep & BASE & "&table=" & TABLE," ","+")
urlk2 = "?lng=" & LNG & "&module=" & request("module") & urlk

' --- IMPORT TABLE ---
if BASE<>"" then
	if TABLE <> "" then%><!-- #include file="../mod/edit/export/import_text.asp"--><%end if
	conn.close
end if

' --- REDIRECTION ---
if SQLstr="" and msg="" and urlk2<>"?"&request.serverVariables("QUERY_STRING") then%>
	<script>location.replace("<%=URL & urlk2%>")</script>
<%end if%>

<html><head><title><%=Version%></title>

	<LINK REL="stylesheet" TYPE="text/css" HREF="styles.css">

</head><body vlink="#7711dd" bgcolor="#FFFFFF"><center>

	<div class=titre><%=trad(0)%></div><br>../<%=baseRep%>
	<%if BASE<>"" then response.write BASE & ".mdb"%><br>
	<%if display<>"" then%><%=trad(42)%> >>
		<A class=a target="_blank" href="<%=replace(display,"\","/")%>"
			><%=replace(display,user_filefolder,"")%></A><br>
	<%end if%>
	<br>

<table border=0 cellspacing=1 cellpadding=0 class=cadre><tr><td>
<table border=0 cellspacing=1 cellpadding=0 bgcolor=#FFFFFF><tr><td>
<table border=0 cellspacing=0 cellpadding=2 class=bgtable>

	<tr><td colspan=3>
			<%=replace(lnglist,"|||",urlk)%>
		</td><td colspan=3><div align=right>
			<A href="http://www.asp-php.net/scripts/asp-php/axs.php"><%=Version%></A>
		</div></td>
	</tr>

<% ' Protection
if not user_OK then%>

	<form method="post"><tr>
		<td align=center colspan=4>Login<BR><input size=<%=SIZEINPUT%> name="login"></td>
		<td align=center>Password<BR><input size=<%=SIZEINPUT%> type="password" name="pass"></td>
		<td align=center><BR><input type="submit" value="OK"></td>
	</tr></form>
	</table>
	</td></tr></table>
	</td></tr></table>

<%else%>

	<script language="JavaScript"><!--
		function test1(fld) {
			// alert if empty field / alerte si champ vide
			if(fld.value=="") { alert("<%=trad(31)%>"); fld.focus(); return false }
			return true
		}
		function test(fld) {
			// alert if bad name / alerte si nom invalide
			var reg = /^[ a-zA-Z0-9._-]+$/
			test0 = (reg.exec(fld.value)!=null)
			if(!test0) { alert("<%=trad(44)%>"); fld.focus() }
			return test0
		}
		function del(type,str) { // confirm deletion / confirme la destruction
			switch(type) {
				case 0:
					if(str.indexOf("DROP TABLE")>-1) msg = "<%=trad(30) & " " & trad(5) & " '" & TABLE & "'"%>"
					else if(str.indexOf("DELETE")>-1) msg = "<%=trad(30) & " " & trad(32)%>"
					else if(str.indexOf("ALTER TABLE")>-1) {
						if(str.indexOf("DROP COLUMN")>-1) msg="<%=trad(30)%>"
						else msg="<%=trad(40)%>"
						msg+= " <%=ucase(trad(6))%>"+str.substr(str.indexOf("COLUMN")+6)
					}
				break
				case 1: msg = "<%=trad(30) & " " & trad(1) & " '" & BASE & "'"%>"; break
			}
			if(confirm(msg+" ?"))
				switch(type) {
					case 0: return true
					default: location.replace(str); break
				}
		}
		function message(n) {
			if(n==1) alert("<%=trad(48)%>")
		}
	//--></script>

	<%' --- DATABASES MANAGEMENT  / GESTION DES BD ---
	%><!--#include file="../mod/edit/bases.asp"--><%

	if BASE<>"" then
		Conn.Open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & basePath%>

		<script language="JavaScript"><!--
			function exporte(type) {
				with(document.myform4)
					if(AxS_strsql.value.toUpperCase().indexOf("SELECT")>-1)
						{ action+="&export="+type; dosql() }
					else
						location.replace("<%=urlk2%>&export="+type)
			}
			function dosql(sql) {
				with(document.myform4) {
					if(sql!=undefined) AxS_strsql.value=sql
					temp = AxS_strsql.value.toUpperCase()
					if((temp.indexOf("DELETE")>-1)
						||((temp.indexOf("DROP TABLE")>-1))
						||((temp.indexOf("ALTER TABLE")>-1))
						) { if(del(0,temp)==true) submit() }
					else submit()
				}
			}
		//--></script>

		<%select case request.queryString("module")%>		
			<%case "online"%>			
				<!--#include file="../mod/online/inc.asp"-->
			<% case else %>
				<!--#include file="../mod/edit/inc.asp"-->
		<%end select
	end if%>

	</table></td><td>
	</td></tr></table>
	<%if BASE<>"" and user_modules and user_plain_base then%></td><td class=bgtable valign=top>
		<!--#include file="../mod/inc.asp"-->
	<%end if%>
	</td></tr></table>

	<!--#include file="../mod/edit/select.asp"-->

<%end if%>
</center></body></html>
<%
' --- ERRORS / ERREURS ---
if err then%><script>alert("<%=err.description%>")</script><%
elseif msg<>"" then%><script>alert("<%=msg%>")</script><%end if

' Display export files / Affiche les fichiers d'export
if display<>"" then%><script>window.open("<%=replace(display,"\","/")%>")</script><%end if
%>