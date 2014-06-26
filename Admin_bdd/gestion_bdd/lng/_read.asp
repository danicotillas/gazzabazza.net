<%' --- READ GLOSSARY / LECTURE DU GLOSSAIRE ---

lng = request("lng")
if lng="" then lng=user_language

Fnm = server.MapPath("../lng/axs_"&lng&".txt")
if FSO.fileExists(Fnm) then
   set inF = FSO.openTextFile(Fnm,1,false)
   trad=split(inF.readAll,VbCrLf,-1,1)
   inF.close
end if

' --- LIST LANGUAGES / LISTE DES LANGUES ---

set fc = FSO.getFolder(server.mapPath(".") & "\..\lng\").files
for each f in fc
	if right(f.name,4) = ".txt" and left(f.name,4) = "axs_" then
		temp = mid(f.name,5,2)
		lnglist = lnglist & "<A href=""" & url & "?lng=" & temp & "|||"">" _
			& "<img src=""../lng/mini" & temp & ".gif"" border=0 width=18 height=12 align=""absmiddle"">" _
			& "</A>" & VbCrLf
	end if
next%>