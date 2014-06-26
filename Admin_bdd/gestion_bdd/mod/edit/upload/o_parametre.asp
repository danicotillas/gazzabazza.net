<!--#include file="../../../_params.asp"-->
<%
TotalMaxSize					= user_upload_maxsize
if user_fileFolder<>"" then user_uploadFolder = server.mapPath("../../"&user_fileFolder) _
else user_uploadFolder = user_baseFolder

'+------------------------------------------------{ Objet : parametre }-----+
'!                                                                          !
'! role : Cet objet contient les diff�rents param�tre pour l'upload de      !
'!        de fichiers sur le serveur.                                       !
'!                                                                          !
'! M�thodes publics :                                                       !
'! ------------------                                                       !
'!  construire(chaine)	 	: constructeur ou chaine est une occurence du     !
'!                          controle filename.                              !
'!  property let repServ(repertoire)                                        !
'!                        : renseigne la donn�e membre p_resServ.           !
'!  property let tailleFichier(taille)                                      !
'!                        : renseigne la donn�e membre p_tailleFichier.     !
'!  property let extensions(ext)                                            !
'!                        : renseigne la donn�e membre p_extensions.        !
'!  property let extToutSauf(ext)                                           !
'!                        : renseigne la donn�e membre p_extToutSauf.       !
'!  estNull()							: true si l'objet est null                        !
'!                        	liste d'objets erreurs.                         !
'!  controleTaille(taille): controle que taille est <= p_tailleFichier ou   !
'!                          taille <= TotalMaxSize.                         !
'!  controleExt(ext)      : controle que ext est present dans p_extensions. !
'!  controleExtSauf(ext)  : controle que ext n est pas dans p_extToutSauf.  !
'!  repertoireServ()      : retourne p_repServ ou Serveur_Repertoire(defaut)!
'!  tailleLimite()        : retourne p_tailleFichier ou TotalMaxSize(defaut)!
'!  AfficheObjet()        : affiche les donn�es membres de l objet          !
'!						                                                              !
'! M�thodes priv�es :                                                       !
'! ------------------                                                       !
'!	sub class_initialize	: constructeur par d�faut. Initialise � null les  !
'!                          donn�es                                         !
'!						                                                              !
'+--------------------------------------------------------------------------+
class parametre
	'// Donn�es membres
	private p_repServ
	private p_tailleFichier
	private	p_extensions
	private p_extToutSauf

	'// M�thodes
	'+---------------------------------------{ M�thode : class_initialize }-----+
	'!                                                                          !
	'! class_initialize                                                         !
	'!                                                                          !
	'! role : Constructeur d'un l'objet parametre � null.                       !
	'!                                                                          !
	'! Parametres :                                                             !
	'!                                                                          !
	'! Valeur retournee :                                                       !
	'!                                                                          !
	'+--------------------------------------------------------------------------+
	private sub class_initialize
		p_repServ					= ""
		p_tailleFichier		= ""
		p_extensions			= ""
		p_extToutSauf			= ""
	end sub
	
	'+---------------------------------------------{ M�thode : construire }-----+
	'!                                                                          !
	'! construire(rep,taille,ext,extSauf)                                       !
	'!                                                                          !
	'! role : Constructeur d'un l'objet.                                        !
	'!                                                                          !
	'! Parametres :                                                             !
	'!                                                                          !
	'! Valeur retournee :                                                       !
	'!                                                                          !
	'+--------------------------------------------------------------------------+
	function construire(rep,taille,ext,extSauf)
		p_repServ					= rep
		p_tailleFichier		= taille
		p_extensions			= ext
		p_extToutSauf			= extSauf
			
	end function
	
	'+-----------------------------------------------{ M�thodes : repServ }-----+
	'!                                                                          !
	'! repServ(repertoire)                                                      !
	'!                                                                          !
	'! role : Renseigne la donn�e membre p_repServ.                             !
	'!                                                                          !
	'! Parametres : repertoire = chaine de caract�re contenant le r�pertoire    !
	'!                           destinataire. exemple : \rep\                  !
	'!                                                                          !
	'! Valeur retournee :                                                       !
	'!                                                                          !
	'+--------------------------------------------------------------------------+
	property let repServ(repertoire)
		p_repServ	= repertoire
	end property

	'+-----------------------------------------{ M�thodes : tailleFichier }-----+
	'!                                                                          !
	'! tailleFichier(taille)                                                    !
	'!                                                                          !
	'! role : Renseigne la donn�e membre p_tailleFichier.                       !
	'!                                                                          !
	'! Parametres : taille = chaine de caract�res contenant la taille limite    !
	'!                       des fichiers � uploader en ko.                     !
	'!                                                                          !
	'! Valeur retournee :                                                       !
	'!                                                                          !
	'+--------------------------------------------------------------------------+
	property let tailleFichier(taille)
		p_tailleFichier	= taille
	end property

	'+--------------------------------------------{ M�thodes : extensions }-----+
	'!                                                                          !
	'! extensions(ext)                                                          !
	'!                                                                          !
	'! role : Renseigne la donn�e membre p_extensions.                          !
	'!                                                                          !
	'! Parametres : ext = chaine de caract�res contenant la liste des extensions!
	'!                    de fichier � uploader sous la forme : "txt htm html". !
	'!                                                                          !
	'! Valeur retournee :                                                       !
	'!                                                                          !
	'+--------------------------------------------------------------------------+
	property let extensions(ext)
		p_extensions	= ext
	end property

	'+-------------------------------------------{ M�thodes : extToutSauf }-----+
	'!                                                                          !
	'! extToutSauf(ext)                                                         !
	'!                                                                          !
	'! role : Renseigne la donn�e membre p_extToutSauf.                         !
	'!                                                                          !
	'! Parametres : ext = chaine de caract�res contenant la liste des extensions!
	'!                    de fichier � ne pas uploader sous la forme :          !
	'!                    "txt htm html".                                       !
	'!                                                                          !
	'! Valeur retournee :                                                       !
	'!                                                                          !
	'+--------------------------------------------------------------------------+
	property let extToutSauf(ext)
		p_extToutSauf		= ext
	end property

	'+------------------------------------------------{ M�thode : estNull }-----+
	'!                                                                          !
	'! estNull()                                                                !
	'!                                                                          !
	'! role : Test si l'objet parametre est un objet null.                      !
	'!                                                                          !
	'! Parametres :                                                             !
	'!                                                                          !
	'! Valeur retournee : true	: l'objet est null                              !
	'!                    false	: l'objet n'est pas null                        !
	'!                                                                          !
	'+--------------------------------------------------------------------------+
	function estNull()

		if	p_repServ	= "" and p_tailleFichier = "" and p_extensions = "" and p_extToutSauf = "" then
				estNull=true
		else
				estNull=false
		end if
	end function

	'+-----------------------------------------{ M�thode : controleTaille }-----+
	'!                                                                          !
	'! controleTaille(taille)                                                   !
	'!                                                                          !
	'! role : Controle si taille est <= � la donn�es membre tailleFichier ou    !
	'!        la constante TotalMaxSize si tailleFichier est � blancs.          !
	'!                                                                          !
	'! Parametres :                                                             !
	'!                                                                          !
	'! Valeur retournee : true	: taille OK                                     !
	'!                    false	: taille KO                                     !
	'!                                                                          !
	'+--------------------------------------------------------------------------+
	function controleTaille(taille)
	
		controleTaille = false
		
		if	p_tailleFichier = "" then
			if taille <= (TotalMaxSize*1000) then controleTaille = true
		else
			if taille <= (p_tailleFichier*1000) then controleTaille = true
		end if

	end function

	'+--------------------------------------------{ M�thode : controleExt }-----+
	'!                                                                          !
	'! controleExt(ext)                                                         !
	'!                                                                          !
	'! role : Controle que ext est dans la liste des extensions autoris�es si   !
	'!        la donn�e membre "p_extensions" est diff�rent de blancs.          !
	'!                                                                          !
	'! Parametres :  ext = extension � controler                                !
	'!                                                                          !
	'! Valeur retournee : true	: controle OK                                   !
	'!                    false	: controle KO                                   !
	'!                                                                          !
	'+--------------------------------------------------------------------------+
	function controleExt(ext)
	
		controleExt = false

		if p_extensions = "" then 
			controleExt = true
		else
			tab=split(p_extensions," ",-1,1)
			for i = 0 to ubound(tab)
				if ext = tab(i) then
					controleExt = true
					exit for
				end if
			next				 
		end if

	end function

	'+----------------------------------------{ M�thode : controleExtSauf }-----+
	'!                                                                          !
	'! controleExtSauf(ext)                                                     !
	'!                                                                          !
	'! role : Controle que ext n est pas dans la liste des extensions � ne pas  !
	'!        uploader si la donn�e membre "p_extToutSauf" est diff�rente de    !
	'!        blancs.                                                           !
	'!                                                                          !
	'! Parametres :  ext = extension � controler                                !
	'!                                                                          !
	'! Valeur retournee : true	: controle OK                                   !
	'!                    false	: controle KO                                   !
	'!                                                                          !
	'+--------------------------------------------------------------------------+
	function controleExtSauf(ext)
	
		controleExtSauf = true
		
		if p_extToutSauf = "" then 
			controleExtSauf = true
		else
			tab=split(p_extToutSauf," ",-1,1)
			for i = 0 to ubound(tab)
				if ext = tab(i) then
					controleExtSauf = false
					exit for
				end if
			next				 
		end if
	end function

	'+-----------------------------------------{ M�thode : repertoireServ }-----+
	'!                                                                          !
	'! repertoireServ()                                                         !
	'!                                                                          !
	'! role : Retourne le r�pertoire destinataire indiqu� dans la donn�e        !
	'!        membre p_repServ ou le r�pertoire par d�faut (constante           !
	'!        Serveur_Repertoire).                                              !
	'!                                                                          !
	'! Parametres :                                                             !
	'!                                                                          !
	'! Valeur retournee :                                                       !
	'!                                                                          !
	'+--------------------------------------------------------------------------+
	function repertoireServ()
	
		if	p_repServ = "" then
			repertoireServ = Serveur_Repertoire 
		else
			repertoireServ = p_repServ
		end if

	end function

	'+-------------------------------------------{ M�thode : tailleLimite }-----+
	'!                                                                          !
	'! tailleLimite()                                                           !
	'!                                                                          !
	'! role : Retourne la taille limite pour l upload indiqu� dans la donn�e    !
	'!        membre p_tailleFichier ou la taille par d�faut (constante         !
	'!        TotalMaxSize).                                              !
	'!                                                                          !
	'! Parametres :                                                             !
	'!                                                                          !
	'! Valeur retournee :                                                       !
	'!                                                                          !
	'+--------------------------------------------------------------------------+
	function tailleLimite()
	
		if	p_tailleFichier = "" then
			tailleLimite = TotalMaxSize 
		else
			tailleLimite = p_tailleFichier
		end if

	end function

	'+-------------------------------------------{ M�thode : AfficheObjet }-----+
	'!                                                                          !
	'! AfficheObjet()                                                           !
	'!                                                                          !
	'! role : Affiche les donn�es membres de l'objet.                           !
	'!                                                                          !
	'! Parametres :                                                             !
	'!                                                                          !
	'! Valeur retournee :                                                       !
	'!                                                                          !
	'+--------------------------------------------------------------------------+
	function AfficheObjet()
		Response.write "********* param�tres ********<br>"
		Response.Write "r�pertoire de destination                 : " & Server.MapPath("\") & repertoireServ() & "<br>"
		Response.write "taille limite des fichiers upload�s	      : " & tailleLimite() & " ko<br>"
		if p_extensions <> "" then Response.write "extensions de fichiers � uploader 	      : " & p_extensions & "<br>"
		if p_extToutSauf <> "" then Response.write "extensions de fichiers � ne pas uploader 	: " & p_extensions & "<br>"
	end function

end class	
%>