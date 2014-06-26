<% ' --- LIST OF FIELDS / LISTE DES CHAMPS --- %>
	<script language="JavaScript"><!--
		function modifysize(fld) {
			with(document.myform3)
				if(test1(sizechamp))
					dosql("ALTER TABLE [<%=TABLE%>] ALTER COLUMN ["+fld+"] TEXT("+sizechamp.value+")") 
		}
		function modifyname(fld,type,size) {
			with(document.myform3)
				if(test(newchamp)) {
					sizechamp.value=size
					action+="&renamefld="+fld+"&typeChamp="+type
					document.myform3.submit()
				}
		}
		function modifytype(fld) {
			with(document.myform3)
				dosql("ALTER TABLE [<%=TABLE%>] ALTER COLUMN ["+fld+"] "+typechamp.value) 
		}
	//--></script>
	
	<tr><td align=center colspan=2><b><%=RS.fields.count & " "%>
		<%if RS.fields.count>1 then response.write trad(7) _
		else response.write trad(6)%></b></td>
		<td colspan=2><i><%=trad(8)%></i></td>
		<td><i><%=trad(9)%></i></td>
		<td><i><%=trad(10)%></i></td>
	</tr>
	<%i=1 : tempSrch=""
	for each Champ in RS.fields
		if Champ.type>=200 and Champ.type<=203 then
			tempSrch = tempSrch & "<option"
			if Champ.name=request.form("fldSrch") then _
				tempSrch = tempSrch & " SELECTED"
			tempSrch = tempSrch & ">" & Champ.name & "</option>"
		end if
	%>
	<tr <%if i mod 2 = 1 then%>class=bgline<%end if%>>
		<td align=center>&nbsp;</td>
		<td align=center><%if RS.fields.count>1 and user_plain_table and left(TABLE,4)<>"AxS_" then%>
			<A href="javascript:dosql('ALTER TABLE [<%=TABLE%>] DROP COLUMN [<%=Champ.name%>]')"
				><%=delimg%></A>
		<%end if%></td>
		<td colspan=2>
			<%if user_plain_table and left(TABLE,4)<>"AxS_" then%>
				<A class=a title="<%=trad(40)%>"
					href="javascript:modifytype('<%=Champ.name%>')"
					><%if Champ.Properties("IsAutoIncrement") then
						isID = Champ.name
						response.write FieldType(-1,true)
					else
						response.write FieldType(Champ.type,true)
					end if%></A>
			<%else
				response.write FieldType(Champ.type,true)
			end if%>
		</td>
		<td>
			<%if user_plain_table and left(TABLE,4)<>"AxS_" and not Champ.Properties("IsAutoIncrement") then%>
				<A title="<%=trad(40)%>" href="javascript:modifyname('<%=Champ.name%>','<%=FieldType(Champ.type,false)%>','<%=Champ.definedsize%>')"><%=Champ.name%></A>
			<%else
				response.write "<b>" & Champ.name & "</b>"
			end if%>
		</td>
		<td><div align=right>
			<%if Champ.type=202 or Champ.type=200 then
				if  user_plain_table and left(TABLE,4)<>"AxS_" then %>
					<A class=a title="<%=trad(40)%>"
						href="javascript:modifysize('<%=Champ.name%>')"><%=Champ.definedsize%></A>
				<%else
					response.write Champ.definedsize
				end if
			end if%>
		&nbsp;</div></td>
	</tr>
	<%i=i+1 : next
	
	if user_plain_table and left(TABLE,4)<>"AxS_" then%>

	<form method="post" name="myform3" onsubmit="return test(this.newchamp)"
		action="<%=URL & urlk2%>">
	<tr>
		<td colspan=2 align=right><input type="submit" value="<%=trad(33)%>"></td>
		<td colspan=2><select name="typechamp">
		<%if isID="" then%><option value="Counter"><%=FieldType(-1,true)%></option><%end if%>
		<%k = array(11,17,2,3,4,5,6,135,202,203,128,205)
			for z = 0 to ubound(k)%>
				<option value="<%=FieldType(k(z),false)%>"
				<%if k(z)=202 then response.write " SELECTED"%>
				><%=FieldType(k(z),true)%></option>
			<%next%>
		</select></td>
		<td><input size=<%=SIZEINPUT%> name="newchamp"></td>
		<td><input name="sizechamp" size=3></td>
	</tr>
	</form>

	<%end if%>
