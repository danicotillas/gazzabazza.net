		<form method="post" name="myform2" action="<%=URL & urlk2%>"
			onsubmit="return test(this.newtable)">
		<tr>
			<td>&nbsp;<b><%=trad(5)%></b></td>
			<td align=center>
				<%if TABLE<>"" and user_plain_table then%>
				<A href="javascript:dosql('DROP TABLE [<%=TABLE%>]')"><%=delimg%></A>
				<%end if%>
			</td>
			<td colspan=2><select name="table" onChange="document.myform2.submit()">

			<% ' --- LIST OF TABLES / LISTE DES TABLES ---
			Set ObjTable = Conn.OpenSchema(20)
			While not ObjTable.eof
				if ObjTable("table_type")="TABLE" then %>
					<option <%if TABLE=ObjTable("table_name") then%>SELECTED<%end if%>>
						<%=ObjTable("table_name")%></option>
				<%end if
				ObjTable.moveNext
			wend
			ObjTable.close%>
			</select></td>
			<td>
				<%if user_plain_table then%><input size=<%=SIZEINPUT%> name="newtable"><%end if%>
			</td>
			<td>
				<%if user_plain_table then%><input type="submit" value="<%=trad(33)%>"><%end if%>
			</td>
		</tr>
		</form>

		<tr>
			<td align=center colspan=6><%if TABLE <> "" then%><table border=0 cellspacing=0 cellpadding=2 class=bgline><tr>
				<td><b class=i2>&nbsp;<%=trad(42)%></b></td>
				<td><A class=a href="javascript:exporte('Text')" TITLE="<%=trad(27)%>"><IMG SRC="../mod/edit/export/txt.gif" WIDTH=13 HEIGHT=16 BORDER=0></A></td>
				<td><A class=a href="javascript:exporte('Excel')" TITLE="Excel (.dif)"><IMG SRC="../mod/edit/export/slk.gif" WIDTH=16 HEIGHT=16 BORDER=0></A></td>
				<td><A class=a href="javascript:exporte('MySQL')" TITLE="MySQL (.sql)"><IMG SRC="../mod/edit/export/mysql.gif" WIDTH=57 HEIGHT=15 BORDER=0></A></td>
				<%=form_import%>
			</tr></table><%end if%>
			</td>
		</tr>
