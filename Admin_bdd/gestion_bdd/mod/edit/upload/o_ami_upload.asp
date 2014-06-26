<!--#include file="o_ctrl_filename.asp"-->
<!--#include file="o_parametre.asp"-->
<%
'+-----------------------------------------------{ Objet : ami_upLoad }-----+
'!                                                                          !
'! role : Cet objet gère la copie de fichiers sur le serveur.               !
'!                                                                          !
'! Méthodes publics :                                                       !
'! ------------------                                                       !
'!  upload(POST,repertoire)                                                 !
'!                  	  	: constructeur                                    !
'!                          (POST= request.binaryread(request.totalbytes))  !
'!                          si repertoire = "" les fichiers sont copiés dans!
'!                          le répertoire par défaut                        !
'!												  (Const Serveur_Repertoire)                      !
'!  repertoireServeur(repertoire)                                           !
'!                        : permet de renseigner un répertoire destinataire !
'!                          sur le serveur autre que celui par défaut.      ! 
'!  tailleFichiersUploades(taille)                                          !
'!                        : permet de renseigner la taille maxi des fichiers!
'!                          uploadés autre que celle par défaut.            ! 
'!  estNull()							: true si l'objet est null                        !
'!	NbreFichiersEcrits()	: retourne le nombre de fichiers écrits.          !
'!	NbreTotalFichiers()		: retourne le nombre total de fichiers transmis.  !
'!  fichiersNonEcrits()		: retourne un tableau contenant tous les fichiers !
'!                          non uploadés.                                   !
'!  fichiersEcrits()			: retourne un tableau contenant tous les fichiers !
'!                          uploadés.                                       !
'!  AfficheObjet          : affiche les données membres de l'objet saul     !
'!                          contenuFichier.                                 !
'!						                                                              !
'! Méthodes privées :                                                       !
'! ------------------                                                       !
'!	sub class_initialize	: constructeur par défaut. Initialise à null les  !
'!                          données                                         !
'!  sub class_terminate		: destructeur - lance la destruction de la donnée !
'!                          membre fichier (classe ctrl_filename)           !
'!  cl_split(chaineBinaire,separateur)                                      !
'!                        : idem commande split sur une chaine binaire.     !
'!						                                                              !
'+--------------------------------------------------------------------------+
class ami_upLoad
	'// Données membres
	public TotalBytes
	public param							'objet paramètre
	public fichier						'tete de la liste des objets ctrl_filename

	'// Méthodes
	'+---------------------------------------{ Méthode : class_initialize }-----+
	'!                                                                          !
	'! class_initialize                                                         !
	'!                                                                          !
	'! role : Constructeur d'un l'objet erreur à null.                          !
	'!                                                                          !
	'! Parametres :                                                             !
	'!                                                                          !
	'! Valeur retournee :                                                       !
	'!                                                                          !
	'+--------------------------------------------------------------------------+
	private sub class_initialize
		TotalBytes					= ""
		set param						= new parametre
		set fichier 				= new ctrl_filename
	end sub
	
	'+----------------------------------------{ Méthode : class_terminate }-----+
	'!                                                                          !
	'! class_terminate                                                          !
	'!                                                                          !
	'! role : Destructeur de l'objet tete de liste et de tous les objets chainés!
	'!                                                                          !
	'! Parametres :                                                             !
	'!                                                                          !
	'! Valeur retournee :                                                       !
	'!                                                                          !
	'+--------------------------------------------------------------------------+
	private sub class_terminate
		set param 	= nothing
		set fichier = nothing
	end sub
	
	'+-------------------------------------------------{ Méthode : upload }-----+
	'!                                                                          !
	'! upload(POST,repertoire)                                                  !
	'!                                                                          !
	'! role : Constructeur de l'objet avec données transmises par un contrôle   !
	'!        Filename. Lance la construction des objets ctrl_filename et écrit !
	'!        dans "repertoire" ou le répertoire par défaut (Serveur_Repertoire)!
	'!        les fichiers transmis.                                            !
	'!                                                                          !
	'! Parametres : POST 			 = Request.BinaryRead(Request.TotalBytes)         !
	'!						  repertoire = repertoire virtuel sur le serveur ou ""        !
	'!                                                                          !
	'! Valeur retournee : true si au moins un fichier est écrit, false sinon.   !
	'!                                                                          !
	'+--------------------------------------------------------------------------+
	function upload(pTotalBytes,HTTP_Content_Type)
		au_moins_un					= false
		TotalBytes					= pTotalBytes

    PosB = InStr(LCase(HTTP_Content_Type), "boundary=") 'Finds boundary
    If PosB > 0 Then separateur = Mid(HTTP_Content_Type, PosB + 9) 'Separateur boundary

		tab=cl_split(TotalBytes,asciiTObinaire(separateur))
		index = 0
		while tab(index) <> ""
			index=index+1
		wend

		if index > 1 then
	  	for i = 0 to index
				set nouveauFichier	= new ctrl_filename
	  		nouveauFichier.construire(tab(i))
				if not nouveauFichier.estNull() then
					au_moins_un = true
		  		set fichier=nouveauFichier.ajoutFilename(fichier)
				else
					set nouveauFichier =nothing
				end if
			next
			if au_moins_un then fichier.EcrireSurServeur(param)
		end if			

		upload = au_moins_un
			
	end function
	
	'+--------------------------------------{ Méthode : repertoireServeur }-----+
	'!                                                                          !
	'! repertoireServeur(repertoire)                                            !
	'!                                                                          !
	'! role : Renseigne le paramètre "repServ" de l'objet donnée membre "param".!
	'!                                                                          !
	'! Parametres : repertoire = chemin du répertoire sous la forme \rep\       !
	'!                                                                          !
	'! Valeur retournee :                                                       !
	'!                                                                          !
	'+--------------------------------------------------------------------------+
	function repertoireServeur(repertoire)
		param.repServ = repertoire
	end function
	
	'+---------------------------------{ Méthode : tailleFichiersUploades }-----+
	'!                                                                          !
	'! tailleFichiersUploades(taille)                                           !
	'!                                                                          !
	'! role : Renseigne le paramètre "tailleFichier" de l'objet donnée membre   !
	'!        "param".                                                          !
	'!                                                                          !
	'! Parametres : taille = valeur numérique en ko                             !
	'!                                                                          !
	'! Valeur retournee :                                                       !
	'!                                                                          !
	'+--------------------------------------------------------------------------+
	function tailleFichiersUploades(taille)
		param.tailleFichier = taille			
	end function
	
	'+-------------------------------------{ Méthode : extensionsUploadee }-----+
	'!                                                                          !
	'! extensionsUploadee(ext)                                                  !
	'!                                                                          !
	'! role : Renseigne le paramètre "extensions" de l'objet donnée membre      !
	'!        "param".                                                          !
	'!                                                                          !
	'! Parametres : ext = chaine de caractères listant les extensions, séparées !
	'!                    par des blancs. exemple : "txt htm html"              !
	'!                                                                          !
	'! Valeur retournee :                                                       !
	'!                                                                          !
	'+--------------------------------------------------------------------------+
	function extensionsUploadee(ext)
		param.extensions = ext		
	end function

	'+----------------------------------{ Méthode : extensionsNonUploadee }-----+
	'!                                                                          !
	'! extensionsNonUploadee(ext)                                               !
	'!                                                                          !
	'! role : Renseigne le paramètre "extToutSauf" de l'objet donnée membre     !
	'!        "param".                                                          !
	'!                                                                          !
	'! Parametres : ext = chaine de caractères listant les extensions, séparées !
	'!                    par des blancs. exemple : "txt htm html"              !
	'!                                                                          !
	'! Valeur retournee :                                                       !
	'!                                                                          !
	'+--------------------------------------------------------------------------+
	function extensionsNonUploadee(ext)
		param.extToutSauf = ext			
	end function

	'+------------------------------------------------{ Méthode : estNull }-----+
	'!                                                                          !
	'! estNull()                                                                !
	'!                                                                          !
	'! role : Test si l'objet ami_upLoad est un objet null.                     !
	'!                                                                          !
	'! Parametres :                                                             !
	'!                                                                          !
	'! Valeur retournee : true	: l'objet est null                              !
	'!                    false	: l'objet n'est pas null                        !
	'!                                                                          !
	'+--------------------------------------------------------------------------+
	function estNull()
		if	fichier.estNull() then
				estNull=true
		else
				estNull=false
		end if
			
	end function

	'+-----------------------------------------------{ Méthode : cl_split }-----+
	'!                                                                          !
	'! cl_split(chaineBinaire,separateur)                                       !
	'!                                                                          !
	'! role : réalise la commande split sur une chaine binaire.                 !
	'!                                                                          !
	'! Parametres : chaineBinaire = chaine binaire à traiter.                   !
	'!              separateur    = chaine binaire à prendre en compte comme    !
	'!                              séparateur.                                 !
	'!                                                                          !
	'! Valeur retournee : tableau                                               !
	'!                                                                          !
	'+--------------------------------------------------------------------------+
	private function cl_split(chaineBinaire,separateur)
		
		dim tab(100000)
		i = 0
	  PosSeparateur = InStrB(1,chaineBinaire, separateur)
	  PosSeparateur_en_cours = PosSeparateur
		while PosSeparateur_en_cours <> 0
		  PosSeparateur = InStrB(PosSeparateur_en_cours+1,chaineBinaire, separateur)
			if PosSeparateur <> 0 then tab(i) = midb(chaineBinaire,PosSeparateur_en_cours+lenb(separateur),PosSeparateur)
			i = i+1
			PosSeparateur_en_cours = PosSeparateur
		wend

		cl_split = tab

	end function

	'+-------------------------------------{ Méthode : NbreFichiersEcrits }-----+
	'!                                                                          !
	'! NbreFichiersEcrits()                                                     !
	'!                                                                          !
	'! role : Retourne le nombre de fichiers écrits.                            !
	'!                                                                          !
	'! Parametres :                                                             !
	'!                                                                          !
	'! Valeur retournee : nombre de fichiers écrits.                            !
	'!                                                                          !
	'+--------------------------------------------------------------------------+
	function NbreFichiersEcrits()
		cpteur = 0

		set fichier_en_cours = fichier
		
		while not fichier_en_cours.estNull()
			if fichier_en_cours.estEcrit() then	cpteur=cpteur+1
			set fichier_en_cours=fichier_en_cours.suivant
		wend
		
		NbreFichiersEcrits=cpteur
		
	end function

	'+--------------------------------------{ Méthode : NbreTotalFichiers }-----+
	'!                                                                          !
	'! NbreTotalFichiers()                                                      !
	'!                                                                          !
	'! role : Nombre de fichier transmis.                                       !
	'!                                                                          !
	'! Parametres :                                                             !
	'!                                                                          !
	'! Valeur retournee : nombre total de fichiers transmis.                    !
	'!                                                                          !
	'+--------------------------------------------------------------------------+
	function NbreTotalFichiers()
		cpteur = 0

		set fichier_en_cours = fichier
		
		while not fichier_en_cours.estNull()
			cpteur=cpteur+1
			set fichier_en_cours=fichier_en_cours.suivant
		wend
		
		NbreTotalFichiers=cpteur
		
	end function

	'+--------------------------------------{ Méthode : fichiersNonEcrits }-----+
	'!                                                                          !
	'! fichiersNonEcrits()                                                      !
	'!                                                                          !
	'! role : Retourne un tableau contenant tous les fichiers non uploadés.     !
	'!                                                                          !
	'! Parametres :                                                             !
	'!                                                                          !
	'! Valeur retournee : tableau contenant les fichier non uploadés.           !
	'!                                                                          !
	'+--------------------------------------------------------------------------+
	function fichiersNonEcrits()
		dim tab(100)
		i = 0
		
		set fichier_en_cours = fichier
		
		while not fichier_en_cours.estNull() and i < 100
			if not fichier_en_cours.estEcrit() then
				tab(i) = fichier_en_cours.repertoire&fichier_en_cours.nom
				i=i+1
			end if
			set fichier_en_cours=fichier_en_cours.suivant
		wend
		
		fichiersNonEcrits=tab
		
	end function

	'+-----------------------------------------{ Méthode : fichiersEcrits }-----+
	'!                                                                          !
	'! fichiersEcrits()                                                         !
	'!                                                                          !
	'! role : Retourne un tableau contenant tous les fichiers uploadés.         !
	'!                                                                          !
	'! Parametres :                                                             !
	'!                                                                          !
	'! Valeur retournee : tableau contenant les fichiers uploadés.              !
	'!                                                                          !
	'+--------------------------------------------------------------------------+
	function fichiersEcrits()
		dim tab(100)
		i = 0
		
		set fichier_en_cours = fichier
		
		while not fichier_en_cours.estNull() and i < 100
			if fichier_en_cours.estEcrit() then
				tab(i) = fichier_en_cours.repertoire&fichier_en_cours.nom
				i=i+1
			end if
			set fichier_en_cours=fichier_en_cours.suivant
		wend
		
		fichiersEcrits=tab
		
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
		Response.write "*****************************<br>"
'		Response.write "TotalBytes 				: " & binaireTOascii(TotalBytes) & "<br>"
		param.AfficheObjet()
		fichier.AfficheObjet()


	end function

end class	
%>