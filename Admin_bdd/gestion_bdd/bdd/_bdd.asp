<%' --- DATABASES MANAGEMENT  / GESTION DES BD 1 ---

baseRep = left(request.queryString("base"),instrrev(request.queryString("base"),"/"))
if request.form("newbdd") <> "" then
	BASE = baseRep & request.form("newbdd")
elseif request.form("baselist") = "../" then
	BASE = left(baseRep,len(baseRep)-1)
	temp = instrrev(BASE,"/")
	if temp > 0 then BASE = left(BASE,temp) else BASE=""
elseif request.form("baselist") = "." then
	BASE = baseRep
elseif request.form("baselist")<>"" then
	BASE = baseRep & request.form("baselist")
else 
	BASE = request.queryString("base")
end if
baseRep = left(BASE,instrrev(BASE,"/"))
BASE = replace(BASE,baseRep,"")

basePath = user_baseFolder & baseRep & BASE & ".mdb"
if not FSO.fileExists(basePath) then BASE = ""

if FSO.folderExists(user_baseFolder) then

	' --- SEARCH blank_xx.mdb ---

	set fc = FSO.getFolder(user_baseFolder).files
	if fc.count>0 then
		for each f in fc
			if right(f.name,4) = ".mdb" and left(f.name,6) = "blank_" then
					vierge = f.name : viergevers = mid(vierge,7,2) : exit for
			end if
		next
	end if

	if vierge = "" then
		msg = "SETUP >> blank_xx.mdb ???"

	elseif request.form("newbdd") <> "" then

		if not FSO.FileExists(basePath) then

			select case request("act")

				' --- COPY DATABASE ---
				case "copy" :
					FSO.CopyFile user_baseFolder & baseRep & request("baselist") & ".mdb",basePath

				' --- RENAME DATABASE ---
				case "rename" :
					FSO.CopyFile user_baseFolder & baseRep & request("baselist") & ".mdb",basePath
					FSO.deleteFile user_baseFolder & baseRep & request("baselist") & ".mdb"

				' --- NEW DATABASE ---
				case "new" :
					FSO.CopyFile user_baseFolder & vierge,basePath

				' --- NEW FOLDER ---
				case "newfld" :
					FSO.CreateFolder(replace(basePath,".mdb",""))
					baseRep = baseRep & request.form("newbdd") & "/"
			
			end select

		end if

		if not err then BASE = request.form("newbdd")

	elseif BASE <> "" then

		' --- COMPACT DATABASE ---
		if request("act") = "compact" then
			Set ObjEngine = Server.CreateObject("JRO.JetEngine")
			BaseSourceStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & basePath & ";"
			BaseDest = user_baseFolder & baseRep & BASE & "_compact.mdb"
			if FSO.FileExists(BaseDest) then FSO.DeleteFile BaseDest
			BaseDestStr = "Provider=Microsoft.Jet.OLEDB.4.0;Jet OLEDB:Engine Type="
			select case vierge
				case "blank_97.mdb" : BaseDestStr = BaseDestStr & "4"
				case "blank_2K.mdb" : BaseDestStr = BaseDestStr & "5"
			end select
			BaseDestStr = BaseDestStr & ";Data Source=" & BaseDest & ";"
			ObjEngine.CompactDatabase BaseSourceStr, BaseDestStr
			Set ObjEngine=Nothing
			if FSO.FileExists(BaseDest) then
				FSO.CopyFile BaseDest,basePath ,True
				FSO.DeleteFile BaseDest
			end if

		' --- DELETE DATABASE ---
		elseif request("act")="del" then
			SQLpath = replace(basePath,".mdb",".sql")
			if FSO.fileExists(SQLpath) then FSO.deleteFile SQLpath
			FSO.deleteFile basePath
		end if

	end if

	basePath = replace(user_baseFolder & baseRep & BASE & ".mdb","/","\")
	if not FSO.fileExists(basePath) then BASE = ""

	' --- LIST OF SUBFOLDERS ---
	formDBa = "<option value="".""></option>"
	if baseRep <> "" then formDBa = formDBa & "<option>../</option>"
	set fc = FSO.getFolder(user_baseFolder & baseRep).subFolders
	for each f in fc
			formDBa = formDBa & "<option "
			formDBa = formDBa & ">" & f.name & "/</option>"
	next
	' --- LIST OF DATABASES ---
	set fc = FSO.getFolder(user_baseFolder & baseRep).files
	for each f in fc
		if right(f.name,4) = ".mdb" and left(f.name,6) <> "blank_" then
			formDBa = formDBa & "<option "
			if BASE & ".mdb" = f.name then formDBa = formDBa & "SELECTED"
			formDBa = formDBa & ">" & replace(f.name,".mdb","") & "</option>"
		end if
	next

else
	msg = "PARAMS >> user_baseFolder <>\n\n" & replace(user_baseFolder,"\","/") & " !"
	user_baseFolder = ""
end if

if BASE <> "" then
	Set Conn = Server.CreateObject("ADODB.Connection")
	Conn.Open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & basePath
	Set RS = Server.CreateObject("ADODB.Recordset")
%>