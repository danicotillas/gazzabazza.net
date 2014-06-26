<%
' --- DISPLAY WITHOUT CODE / AFFICHE SANS CODE ---

function nohtml(txt)
	if not isnull(txt) then
		temp = replace(txt,"&","&amp;")
		temp = replace(temp,"""","&quot;")
		nohtml = replace(temp,"<","&lt;")
	end if
end function

' --- MAKE LINK IF URL IN FIELD / CREE UN LIEN SI LE CHAMP CONTIENT UNE URL

function makeUrl(txt)
	dim regEx,temp
	if not isnull(txt) then
		Set regEx = New RegExp
		' --- test E-mail
		regEx.pattern = "^([a-zA-Z0-9]+(([\.\-\_]?[a-zA-Z0-9]+)+)?)\@(([a-zA-Z0-9]+[\.\-\_])+[a-zA-Z0-9]{2,4})$"
		if regEx.test(txt) then _
			temp = "<A class=a href=""mailto:" & txt & """>" & txt & "</A>"
		' --- test URL
		regEx.pattern = "^http://[\w.%/?=&#:;+-]{4,}$"
		if regEx.test(txt) then _
			temp = "<A class=a target=""_blank"" href=""" & txt & """>" & txt & "</A>"
		if temp = "" then makeUrl = nohtml(txt) else makeUrl = temp
	end if
end function


if BASE<>"" then

	' --- IF SELECT REQUEST / SI REQUETE SELECT ---
	if TABLE<>"" AND instr(lcase(SQLstr),"select ")>0 then
		RS.open SQLstr,conn,3,3 %>
		<A name="listing"><br></A>
		<%response.write SQLstr&"<BR><BR>"

		' --- EDITION IF 1 RECORD / SI UNE FICHE ---
		if RS.recordcount = 1 then
			function min(a,b)
				if a>b then min=b else min=a
			end function%>
			<table class=bgtable>
			<form method="post" name="myform5" action="<%=URL & urlk2%>">			
				<input type="hidden" name="AxS_strsql" value="<%=SQLstr%>">
			<%k=1 : isID=false
			for each Champ in RS.fields%>
				<tr>
					<td><b><%=Champ.name%></b><br>
						<i><% if Champ.Properties("IsAutoIncrement") then
							isID=true
							response.write fieldType(-1,true) _
								& "<input type='hidden' name='idname' value='" & Champ.name & "'>"
						else
							response.write fieldType(Champ.type,true)
						end if
						if Champ.type=200 or Champ.type=202 then _
							response.write " " & Champ.definedsize%></i>
					</td><td>
						<%if Champ.Properties("IsAutoIncrement") then
							response.write RS(Champ.name) _
								& "<input type='hidden' name='idval' value='" & RS(Champ.name) & "'>"
						else
							select case Champ.type
								case 2,3,4,5,6,17,135 : ' num & date
									response.write "<input size=" & SIZEINPUT/2 & " name='" & Champ.name _
										& "' value='" & RS(Champ.name) & "'>"
								case 11 : ' boolean
									temp = "<select name='" & Champ.name & "'><option"
									if(RS(Champ.name)) then temp = temp & " SELECTED"
									temp = temp & ">True</option><option"
									if not(RS(Champ.name)) then temp = temp & " SELECTED"
									response.write temp & ">False</option></select>"
								case 200,202 : ' text
									response.write "<input maxlength=" & Champ.definedsize & " size=" & min(SIZEINPUT*5,Champ.definedsize) _
									& " name=""" & Champ.name & """ value=""" & nohtml(RS(Champ.name)) & """>"
								case 201,203 : ' memo
									response.write "<textarea rows=3 cols=" & SIZEINPUT*5 & " name='" _
									& Champ.name & "'>" & nohtml(RS(Champ.name)) & "</textarea>"
							end select
						end if%>
					</td></tr>
				<%k=k+1 : next
				if isID then%>
					<tr><td align=center colspan=2><input type="submit" name="submit" value="<%=trad(38)%>">&nbsp;&nbsp;&nbsp;<input type="submit" name="submit" value="<%=trad(49)%>"></td></tr>
				<%end if%>
				</form>
				</table>

			<% ' --- LISTING IF MANY RECORDS / SI PLUSIEURS FICHES ---
			elseif RS.recordcount>0 then
				strnav=""
				for k=1 to RS.recordCount step user_nbrecord
					temp = int(k/user_nbrecord)+1
					if request("record") = cStr(k) or (request("record")="" and k=1) then _
						strnav = strnav & "&nbsp;<font color=red><b>" & temp & "</b></font>&nbsp; " _
					else _
						strnav = strnav & "<A href='javascript:gorecord(" & k & ")'>&nbsp;" & temp & "&nbsp;</A> "
				next
				response.write "<hr size=1><b>" & RS.recordCount & " " & trad(15) & "</b> : " & trad(45) & " " & strnav & "<hr size=1><br>"%>

				<script language="JavaScript"><!--
					function gorecord(x) {
						with(document.doboxdel) {
							action+="&record="+x
							submit()
						}
					}
				//--></script>

				<table class=bgtable>
					<tr class=bgline><td></td><td align=center><A href="javascript:document.doboxdel.submit()"><%=delimg%></A></td>
					<%isID = ""
					for each Champ in RS.fields
						if Champ.Properties("IsAutoIncrement") then isID=Champ.name%>
						<td align=center><b><%=Champ.name%></b></td>
					<%next%>
					</tr>
					<form name="doboxdel" method="post" action="<%=URL & urlk2%>">
						<input type="hidden" name="AxS_strsql" value="<%=SQLstr%>">
						<input type="hidden" name="IDfld" value="<%=isID%>">
						<input type="hidden" name="record" value="<%=request("record")%>">
				<% k=1 : rec0 = 1 : if request("record")<>"" then rec0=cInt(request("record"))
				while not RS.eof and k<rec0+user_nbrecord
					if k>=rec0 then%>
						<tr <%if k mod 2 = 0 then%>class=bgline<%end if%>>
							<td valign=top><%if isID<>"" then%>
								<A href="javascript:dosql('SELECT * FROM [<%=TABLE%>] WHERE [<%=isID & "]=" & RS(isID)%>')"
								><%=trad(37)%></A>
							<%end if%></td><td valign=top><%if isID<>"" then%>
								<input type=checkbox name="boxdel" value="<%=RS(isID)%>">
							<%end if%></td>
							<%for each Champ in RS.fields%><td valign=top>
								<%select case Champ.type
									case 2,3,4,5,6,17,135 : ' num & date
										response.write "<div align=right>" & RS(Champ.name) & "</div>"
									case 11 : ' boolean
										temp = "<input type='checkbox' DISABLED"
										if(RS(Champ.name)) then temp = temp & " CHECKED"
										response.write temp & ">"
									case 200,202 : ' text
										response.write makeUrl(RS(Champ.name))
									case 201,203 : ' memo
										response.write nohtml(left(RS(Champ.name),255))
								end select%>
							</td><%next%>
						</tr>
					<%end if
					k = k+1 : RS.moveNext
				wend%>
				</form>
				</table><br>

			<%end if%>
		<script>location.replace("#listing")</script>

		<% RS.close
	end if
	Conn.close
end if
%>