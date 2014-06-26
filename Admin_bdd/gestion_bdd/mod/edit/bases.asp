<%' --- DATABASES MANAGEMENT  / GESTION DES BD 2 ---

if user_baseFolder <> "" then%>

   <SCRIPT language="JavaScript"><!--
      function newfld() {
         with(document.formbdd)
            if(test(newbdd)) { action+="&act=newfld"; submit() }
      }
      function newDB() {
         with(document.formbdd)
            if(test(newbdd)) { action+="&act=new"; submit() }
      }
      function copyDB() {
         with(document.formbdd)
            if(test(newbdd)) { action+="&act=copy"; submit() }
      }
      function renameDB() {
         with(document.formbdd)
            if(test(newbdd)) { action+="&act=rename"; submit() }
      }
	//--></SCRIPT>

<form method="post" name="formbdd" action="<%=URL & urlk2%>">
	<tr><td>&nbsp;<b><%=trad(1)%></b>&nbsp;</td>
		<th>
			<%if BASE<>"" and user_plain_base then%><A href="javascript:del(1,'<%=urlk2%>&act=del')"><%=delimg%></A><%end if%><BR>
		</th><td>
			<select name="baselist" onChange="formbdd.submit()">
				<%=formDBa%>
			</select><BR>
		</td><td align=right>
			<%if user_plain_base then%>
			<A TITLE="<%=trad(43)%>" href="javascript:newfld()"
				><IMG SRC="../bdd/mininewfld.gif" WIDTH="16" HEIGHT="15" BORDER=0></A>
			<%end if%>
		</td><th>
			<%if vierge <>"" and user_plain_base then%><input size=<%=SIZEINPUT%> name="newbdd"><%end if%><BR>
		</th><td>
			<%if vierge <>"" and user_plain_base then%><input type="button" onclick="newDB()" value="<%=trad(33)%>"><%end if%><BR>
	<%if BASE <> "" then
		' --- Size of database --- %>
		</td></tr><tr><th colspan=2>
			<%temp = int(FSO.getFile(basePath).size/1024) : temp2 = "Ko"
			if temp > 1024 then
				temp = formatnumber(temp/1024,2) : temp2 = "Mo"
			end if
			response.write temp & " " & temp2%>
		</th><th colspan=4>
			<%if vierge <> "" and user_plain_base then
				' --- Actions on database --- %>
				<A href="<%=urlk2%>&act=compact"><%=trad(2) & " " & viergevers%></A> -
				<A href="javascript:renameDB()"><%=trad(3)%></A> -
				<A href="javascript:copyDB()"><%=trad(4)%></A>
		<%end if
	end if%>
	</th></tr>
</form>
<%end if%>