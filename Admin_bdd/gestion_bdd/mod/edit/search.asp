			<tr><td colspan=3 nowrap>
				<A href="javascript:dosql('DELETE * FROM [<%=TABLE%>]')"><%=delimg%></A>
				<A href="javascript:dosql('SELECT * FROM [<%=TABLE%>]')" title="<%=trad(12)%>"><%=RS.recordcount & " "%>
				<%if RS.recordcount > 1 then response.write trad(15) _
				else response.write trad(14)%></A>
				</td><td colspan=3 nowrap>
				<%if isID<>"" then%>
					<A href="<%=url&urlk2&"&newfi="&isID%>"><%=trad(39)%></A>
				<%end if%>
				</td>
			</tr>
			</form>
	
			<% ' Search Form / Form de Recherche
			if tempSrch <> "" and RS.recordcount>1 then %>
			<tr>
				<td colspan=6 align=center>
					<table class=bgline border=0 cellspacing=0 cellpadding=2>
					<form method="post" action="<%=URL & urlk2%>"><tr>
						<td><A href="javascript:message(1)"><IMG SRC="miniaide.gif" WIDTH="16" HEIGHT="16" BORDER=0></A></td>
						<td><b><%=trad(46)%></b></td>
						<td><input size=<%=SIZEINPUT*3/4%> name="search" value="<%=request.form("search")%>"></td>
						<td><%=trad(47)%></td>
						<td><select name="fldSrch"><%=tempSrch%></select></td>
						<td><input type="submit" value="<%=trad(34)%>">
							<input type="hidden" name="idSrch" value="<%=isID%>"></td>
					</tr></form></table>
				</td>
			</tr>
			<% end if %>
