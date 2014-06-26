<% ' Import TEXT

if user_fileFolder<>"" then import_path = server.mapPath(user_fileFolder) _
else import_path = user_baseFolder

function okFile(Fnm)
	dim temp,temp2,i
	okFile = true
	if right(Fnm,4)=".txt" then
		set inF = FSO.openTextFile(import_path & "\" & Fnm)
		temp = split(inF.readAll,sep2,-1,1)
		inF.close
		' Liste des champs
		temp2 = split(temp(0),sep1,-1,1)
      RS.Open "SELECT * FROM [" & TABLE & "]",Conn,3,3
		' Les mêmes que la table ?
		if RS.fields.count <> ubound(temp2) then
			okFile = false
		else
			for each Champ in RS.fields
				for i = 0 to ubound(temp2)
					if Champ.name = temp2(i) then exit for
				next
				if i > ubound(temp2) then
					okFile = false
					exit for
				end if
			next
		end if
		RS.close
	else
		okFile = false
	end if
end function

FORM_IMPORT = "<form name=""import"" method=""post"" action=""" & URL & urlk2 & """>" _
	& "<td>&nbsp;<A href=""javascript:void(0)"" onclick=""window.open('../mod/edit/upload/upload.htm','','scrollbars=yes,resizable=yes,width=370,height=145,left=20,top=20')"">Import</A></td>"
	tmp = "<td><select class=f7 name=""importxt"">"
set fc = FSO.getFolder(import_path).files : i=0
for each f in fc
	if okFile(f.name) then
		i = i+1
		tmp = tmp & "<option>" & replace(f.name,".txt","") & "</option>"
	end if
next
if i > 0 then form_import = form_import & tmp & "</select></td>" _
	& "<td><input type=""Submit"" value=""OK""></td></form>" _
else form_import = form_import & "</form>"

' === IMPORT ===

if request.form("importxt") <> "" then
	set inF = FSO.openTextFile(import_path & "\" & request.form("importxt") & ".txt")
	temp = split(inF.readAll,sep2,-1,1)
	inF.close
	' Liste des champs
	temp2 = split(temp(0),sep1,-1,1)
	RS.Open "SELECT TOP 1 * FROM [" & TABLE & "]",Conn,3,3
	' pour chaque ligne
	for i = 1 to ubound(temp)
		temp3 = split(temp(i),sep1,-1,1)
		' Vérification du nombre de champs
		if ubound(temp3) = ubound(temp2) then
			RS.addNew
			for j = 0 to ubound(temp2)
				for each Champ in RS.fields
					if Champ.name = temp2(j) then
						' Pas l'auto-increment
						if not Champ.Properties("IsAutoIncrement") and temp3(j)<>"" then
							select case Champ.type
								case 11 : RS(temp2(j)) = (lcase(temp3(j))="vrai" OR lcase(temp3(j))="true")
								case else : RS(temp2(j)) = temp3(j)
							end select
						end if
						exit for
					end if
				next
			next
			RS.update
		end if
	next
	RS.close
end if

%>