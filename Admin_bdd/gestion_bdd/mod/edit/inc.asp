		<tr><td align=center colspan=6><hr size=1 color=#000088></td>

				<!--#include file="tables.asp"-->

				<%if TABLE<>"" then
					RS.Open "SELECT * FROM [" & TABLE & "]",Conn,3,3 %>

					<tr><td align=center colspan=6><hr size=1 color=#000088></td>

					<!--#include file="fields.asp"-->

					<tr><td align=center colspan=6><hr size=1 color=#000088></td></tr>

					<!--#include file="sql.asp"-->

					<tr><td align=center colspan=6><hr size=1 color=#000088></td></tr>

					<!--#include file="search.asp"-->

					<% RS.close
				end if%>
