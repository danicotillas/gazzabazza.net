<%
	if SQLstr <> "" then
		' --- EXECUTE SQL IF NOT SELECT / SI PAS SELECT ---
		if instr(lcase(SQLstr),"select ")<1 then conn.execute(SQLstr)

		' --- SAVE SQL / ENREGISTRE LA REQUETE ---
		if request("act") = "savesql" then
			temp = ""
			set inF = FSO.openTextFile(SQLpath,1,true)
			while not inF.atEndOfStream
				temp2 = ucase(inF.readLine)
				if temp2 <> ucase(SQLstr) then temp = temp & temp2 & VbCrLf
			wend
			inF.close
			temp = SQLstr & VbCrLf & temp
			set inF = FSO.openTextFile(SQLpath,2,true)
			inF.write(temp)
			inF.close

		' --- DELETE SQL / DETRUIT LA REQUETE
		elseif request("act") = "delsql" and FSO.fileExists(SQLpath) then
			set inF = FSO.openTextFile(SQLpath,1,false)
			temp = replace(inF.readAll,SQLstr & VbCrLf,"")
			inF.close
			if temp = "" then
				FSO.DeleteFile SQLpath
			else					
				set inF = FSO.openTextFile(SQLpath,2,false)
				inF.write(temp)
				inF.close
			end if
		end if

	end if
%>