<%			if instr(lcase(SQLstr),"select ")>0 then SQL = SQLstr _
			else SQL = "SELECT * FROM [" & TABLE & "]"
			RS.open SQL,conn,3,3
			select case request("export")
				case "Text": ext = ".txt"
				case "Excel": ext = ".dif"
				case "MySQL": ext = ".sql"
			end select
			if user_fileFolder<>"" then
				Fnm = server.mapPath(user_fileFolder) & "\" & BASE & "_" & TABLE & ext
				set inF = FSO.openTextFile(Fnm,2,true)
			else
				Fnm = user_baseFolder & baseRep & BASE & "_" & TABLE & ext
				set inF = FSO.openTextFile(Fnm,2,true)
			end if
			select case request("export")
				case "Text":
					%><!-- #include file="export_text.asp"--><%
				case "Excel":
					%><!-- #include file="export_dif.asp"--><%
				case "MySQL":
					%><!-- #include file="export_mysql.asp"--><%
			end select
			RS.close
			inF.close
			if user_fileFolder<>"" then
				display = user_fileFolder & BASE & "_" & TABLE & ext
				msg = ucase(trad(42)) & " >> " & replace(display,"\","/")
			else
				msg = ucase(trad(42)) & " >> ../" & mid(Fnm,instrrev(Fnm,"\")+1)
			end if
%>