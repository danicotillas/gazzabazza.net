<%' --- USER PARAMS / PARAMETRES UTILISATEUR ---

' default language (please, send me your translation of lng/AxS_xx.txt if so !)
user_language = "fr" ' or "en" , "de"

user_modules = true ' Show module menu if user_plain_base = true ?
user_plain_base = true ' plain or light version for DATABASES ?
user_plain_table = true ' plain or light version for TABLES ?

' DataBase folder : Called by elt/admin_access.asp (+1 ..\)
'user_baseFolder = server.mapPath("../try/database") & "\"
user_baseFolder = server.mapPath("../../../magia") & "\"
 
'user_baseFolder = server.mapPath("../../../espanol/Forum_es/dbase") & "\"
'user_baseFolder = server.mapPath("../../../english/Forum_en/dbase") & "\"
'user_baseFolder = server.mapPath("../../francais/Forum_fr/dbase") & "\"


' user_baseFolder = server.mapPath("..\..\..\..\database") & "\"

' Number of records per page in SELECT
user_nbrecord = 20 

' --- PROTECTION ---
' let empty if no protection required
user_login = ""
user_pass = ""

' --- EXPORT ---
' write unprotected folder for export files
user_fileFolder = "../try/export/" ' Called by elt/admin_access.asp (+1 ..\) or empty if user_BaseFolder used
user_upload_maxsize = 100 ' Ko
' Separators for Text export files (sepXb to replace in fields)
sep1 = chr(9) : sep1b = "   "
sep2 = VbCrLf : sep2b = chr(13)
' increase if needed (export of large tables)
Server.ScriptTimeout = 5*60 ' seconds

' --- DEBUG ---
user_debug = false ' true to try to debug if needed...:o)
%>