<!--#include file="../../../_params.asp"-->
<%
TotalMaxSize					= user_upload_maxsize
if user_fileFolder<>"" then user_uploadFolder = server.mapPath("../../"&user_fileFolder) _
else user_uploadFolder = user_baseFolder

'+------------------------------------------------{ Objet : parametre }-----+
'!                                                                          !
'! role : Cet objet contient les différents paramètre pour l'upload de      !
'!        de fichiers sur le serveur.                                       !
'!                                                                          !
'! Méthodes publics :                                                       !
'! ------------------                                                       !
'!  construire(chaine)	 	: constructeur ou chaine est une occurence du     !
'!                          controle filename.                              !
'!  property let repServ(repertoire)                                        !
'!                        : renseigne la donnée membre p_resServ.           !
'!  property let tailleFichier(taille)                                      !
'!                        : renseigne la donnée membre p_tailleFichier.     !
'!  property let extensions(ext)                                            !
'!                        : renseigne la donnée membre p_extensions.        !
'!  property let extToutSauf(ext)                                           !
'!                        : renseigne la donnée membre p_extToutSauf.       !
'!  estNull()							: true si l'objet est null                        !
'!                        	liste d'objets erreurs.                         !
'!  controleTaille(taille): controle que taille est <= p_tailleFichier ou   !
'!                          taille <= TotalMaxSize.                         !
'!  controleExt(ext)      : controle que ext est present dans p_extensions. !
'!  controleExtSauf(ext)  : controle que ext n est pas dans p_extToutSauf.  !
'!  repertoireServ()      : retourne p_repServ ou Serveur_Repertoire(defaut)!
'!  tailleLimite()        : retourne p_tailleFichier ou TotalMaxSize(defaut)!
'!  AfficheObjet()        : affiche les données membres de l objet          !
'!						                                                              !
'! Méthodes privées :                                                       !
'! ------------------                                                       !
'!	sub class_initialize	: constructeur par défaut. Initialise à null les  !
'!                          données                                         !
'!						                                                              !
'+--------------------------------------------------------------------------+
class parametre
	'// Données membres
	private p_repServ
	private p_tailleFichier
	private	p_extensions
	private p_extToutSauf

	'// Méthodes
	'+---------------------------------------{ Méthode : class_initialize }-----+
	'!                                                                          !
	'! class_initialize                                                         !
	'!                                                                          !
	'! role : Constructeur d'un l'objet parametre à null.                       !
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
	
	'+---------------------------------------------{ Méthode : construire }-----+
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
	
	'+-----------------------------------------------{ Méthodes : repServ }-----+
	'!                                                                          !
	'! repServ(repertoire)                                                      !
	'!                                                                          !
	'! role : Renseigne la donnée membre p_repServ.                             !
	'!                                                                          !
	'! Parametres : repertoire = chaine de caractère contenant le répertoire    !
	'!                           destinataire. exemple : \rep\                  !
	'!                                                                          !
	'! Valeur retournee :                                                       !
	'!                                                                          !
	'+--------------------------------------------------------------------------+
	property let repServ(repertoire)
		p_repServ	= repertoire
	end property

	'+-----------------------------------------{ Méthodes : tailleFichier }-----+
	'!                                                                          !
	'! tailleFichier(taille)                                                    !
	'!                                                                          !
	'! role : Renseigne la donnée membre p_tailleFichier.                       !
	'!                                                                          !
	'! Parametres : taille = chaine de caractères contenant la taille limite    !
	'!                       des fichiers à uploader en ko.                     !
	'!                                                                          !
	'! Valeur retournee :                                                       !
	'!                                                                          !
	'+--------------------------------------------------------------------------+
	property let tailleFichier(taille)
		p_tailleFichier	= taille
	end property

	'+--------------------------------------------{ Méthodes : extensions }-----+
	'!                                                                          !
	'! extensions(ext)                                                          !
	'!                                                                          !
	'! role : Renseigne la donnée membre p_extensions.                          !
	'!                                                                          !
	'! Parametres : ext = chaine de caractères contenant la liste des extensions!
	'!                    de fichier à uploader sous la forme : "txt htm html". !
	'!                                                                          !
	'! Valeur retournee :                                                       !
	'!                                                                          !
	'+--------------------------------------------------------------------------+
	property let extensions(ext)
		p_extensions	= ext
	end property

	'+-------------------------------------------{ Méthodes : extToutSauf }-----+
	'!                                                                          !
	'! extToutSauf(ext)                                                         !
	'!                                                                          !
	'! role : Renseigne la donnée membre p_extToutSauf.                         !
	'!                                                                          !
	'! Parametres : ext = chaine de caractères contenant la liste des extensions!
	'!                    de fichier à ne pas uploader sous la forme :          !
	'!                    "txt htm html".                                       !
	'!                                                                          !
	'! Valeur retournee :                                                       !
	'!                                                                          !
	'+--------------------------------------------------------------------------+
	property let extToutSauf(ext)
		p_extToutSauf		= ext
	end property

	'+------------------------------------------------{ Méthode : estNull }-----+
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

	'+-----------------------------------------{ Méthode : controleTaille }-----+
	'!                                                                          !
	'! controleTaille(taille)                                                   !
	'!                                                                          !
	'! role : Controle si taille est <= à la données membre tailleFichier ou    !
	'!        la constante TotalMaxSize si tailleFichier est à blancs.          !
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

	'+--------------------------------------------{ Méthode : controleExt }-----+
	'!                                                                          !
	'! controleExt(ext)                                                         !
	'!                                                                          !
	'! role : Controle que ext est dans la liste des extensions autorisées si   !
	'!        la donnée membre "p_extensions" est différent de blancs.          !
	'!                                                                          !
	'! Parametres :  ext = extension à controler                                !
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

	'+----------------------------------------{ Méthode : controleExtSauf }-----+
	'!                                                                          !
	'! controleExtSauf(ext)                                                     !
	'!                                                                          !
	'! role : Controle que ext n est pas dans la liste des extensions à ne pas  !
	'!        uploader si la donnée membre "p_extToutSauf" est différente de    !
	'!        blancs.                                                           !
	'!                                                                          !
	'! Parametres :  ext = extension à controler                                !
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

	'+-----------------------------------------{ Méthode : repertoireServ }-----+
	'!                                                                          !
	'! repertoireServ()                                                         !
	'!                                                                          !
	'! role : Retourne le répertoire destinataire indiqué dans la donnée        !
	'!        membre p_repServ ou le répertoire par défaut (constante           !
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

	'+-------------------------------------------{ Méthode : tailleLimite }-----+
	'!                                                                          !
	'! tailleLimite()                                                           !
	'!                                                                          !
	'! role : Retourne la taille limite pour l upload indiqué dans la donnée    !
	'!        membre p_tailleFichier ou la taille par défaut (constante         !
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

	'+-------------------------------------------{ Méthode : AfficheObjet }-----+
	'!                                                                          !
	'! AfficheObjet()                                                           !
	'!                                                                          !
	'! role : Affiche les données membres de l'objet.                           !
	'!                                                                          !
	'! Parametres :                                                             !
	'!                                                                          !
	'! Valeur retournee :                                                       !
	'!                                                                          !
	'+--------------------------------------------------------------------------+
	function AfficheObjet()
		Response.write "********* paramètres ********<br>"
		Response.Write "répertoire de destination                 : " & Server.MapPath("\") & repertoireServ() & "<br>"
		Response.write "taille limite des fichiers uploadés	      : " & tailleLimite() & " ko<br>"
		if p_extensions <> "" then Response.write "extensions de fichiers à uploader 	      : " & p_extensions & "<br>"
		if p_extToutSauf <> "" then Response.write "extensions de fichiers à ne pas uploader 	: " & p_extensions & "<br>"
	end function

end class	
%>