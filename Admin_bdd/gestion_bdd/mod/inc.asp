<table><tr><td align=center><b>Modules</b></td></tr>
	<tr><td align=center <%
			if request("module")="edit" or request("module")="" then response.write "bgcolor=#FFFFFF" else response.write "class=bgline"%>>
			<A href="<%=URL & "?lng=" & LNG & "&module=edit" & urlk%>">Edition</A></td></tr>
	<tr><td align=center <%
			if request("module")="online" then response.write "bgcolor=#FFFFFF" else response.write "class=bgline"%>>
			<A href="<%=URL & "?lng=" & LNG & "&module=online" & urlk%>">Online</A></td></tr>
</table>
