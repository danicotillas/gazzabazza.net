	<script language="JavaScript"><!--
		function savesql() {
			with(document.myform4)
				if(test1(AxS_strsql))
					{ action+="&act=savesql"; dosql() }
		}
		function delsql() {
			with(document.myform4)
				if(listsql.options[listsql.selectedIndex].value!="")
					{ action+="&act=delsql"; dosql() }
		}
		function readsql() {
			with(document.myform4)
				dosql(listsql.options[listsql.selectedIndex].value)
		}
	--></script>

		<form method="post" name="myform4" action="<%=URL & urlk2%>">
		<tr>
			<td align=center><b><%=trad(11)%></b></td>
			<td align="right"><A TITLE="<%=trad(41)%>" href="javascript:savesql()"
				><IMG SRC="minisave.gif" WIDTH=14 HEIGHT=14 BORDER=0></td>
			<td align=center colspan=3><textarea wrap=virtual rows=2 cols=<%=SIZEINPUT*2-3%> name="AxS_strsql"><%=SQLstr0%></textarea></td>
			<td align=center><input type="button" value="<%=trad(34)%>" onclick="dosql()"></td>
		</tr>
		<tr>
			<td align=right><A href="javascript:delsql()"
				><%=delimg%></td>
			<td colspan=5>
				<select class=f7 style="width:300"
				name="listsql" onchange="readsql()">
			<% ' LIST OF REQUESTS / LISTE DES REQUETES
			if FSO.fileExists(SQLpath) then%>
				<option></option>
				<%set inF = FSO.openTextFile(SQLpath,1,false)
				while not inF.atEndOfStream
					temp = inF.readLine%>
					<option <%if temp = SQLstr then response.write "SELECTED"%>
						value="<%=temp%>"><%=temp%></option>
				<%wend
				inF.close
				end if%>
				<option <%if SQLstr="" then response.write "SELECTED"%>></option>
				<option value="<%temp = "SELECT * FROM [" & TABLE & "]"
					if isID<>"" then temp = temp & " ORDER BY [" & isID & "] DESC"
					response.write temp%>"
					<%if SQLstr = temp then response.write "SELECTED"%>><%=trad(12)%></option>
				<option value="<%temp = "DELETE * FROM [" & TABLE & "]"
					response.write temp%>"><%=trad(13)%></option>
			</select></td>
		</tr>
