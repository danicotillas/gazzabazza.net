<!--#include file="o_ctrl_filename.asp"-->
<!--#include file="o_parametre.asp"-->
<%
'+-----------------------------------------------{ Objet : ami_upLoad }-----+
'!                                                                          !
'! role : Cet objet g�re la copie de fichiers sur le serveur.               !
'!                                                                          !
'! M�thodes publics :                                                       !
'! ------------------                                                       !
'!  upload(POST,repertoire)                                                 !
'!                  	  	: constructeur                                    !
'!                          (POST= request.binaryread(request.totalbytes))  !
'!                          si repertoire = "" les fichiers sont copi�s dans!
'!                          le r�pertoire par d�faut                        !
'!												  (Const Serveur_Repertoire)                      !
'!  repertoireServeur(repertoire)                                           !
'!                        : permet de renseigner un r�pertoire destinataire !
'!                          sur le serveur autre que celui par d�faut.      ! 
'!  tailleFichiersUploades(taille)                                          !
'!                        : permet de renseigner la taille maxi des fichiers!
'!                          upload�s autre que celle par d�faut.            ! 
'!  estNull()							: true si l'objet est null                        !
'!	NbreFichiersEcrits()	: retourne le nombre de fichiers �crits.          !
'!	NbreTotalFichiers()		: retourne le nombre total de fichiers transmis.  !
'!  fichiersNonEcrits()		: retourne un tableau contenant tous les fichiers !
'!                          non upload�s.                                   !
'!  fichiersEcrits()			: retourne un tableau contenant tous les fichiers !
'!                          upload�s.                                       !
'!  AfficheObjet          : affiche les donn�es membres de l'objet saul     !
'!                          contenuFichier.                                 !
'!						                                                              !
'! M�thodes priv�es :                                                       !
'! ------------------                                                       !
'!	sub class_initialize	: constructeur par d�faut. Initialise � null les  !
'!                          donn�es                                         !
'!  sub class_terminate		: destructeur - lance la destruction de la donn�e !
'!                          membre fichier (classe ctrl_filename)           !
'!  cl_split(chaineBinaire,separateur)                                      !
'!                        : idem commande split sur une chaine binaire.     !
'!						                                                              !
'+--------------------------------------------------------------------------+
class ami_upLoad
	'// Donn�es membres
	public TotalBytes
	public param							'objet param�tre
	public fichier						'tete de la liste des objets ctrl_filename

	'// M�thodes
	'+---------------------------------------{ M�thode : class_initialize }-----+
	'!                                                                          !
	'! class_initialize                                                         !
	'!                                                                          !
	'! role : Constructeur d'un l'objet erreur � null.                          !
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
	
	'+----------------------------------------{ M�thode : class_terminate }-----+
	'!                                                                          !
	'! class_terminate                                                          !
	'!                                                                          !
	'! role : Destructeur de l'objet tete de liste et de tous les objets chain�s!
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
	
	'+-------------------------------------------------{ M�thode : upload }-----+
	'!                                                                          !
	'! upload(POST,repertoire)                                                  !
	'!                                                                          !
	'! role : Constructeur de l'objet avec donn�es transmises par un contr�le   !
	'!        Filename. Lance la construction des objets ctrl_filename et �crit !
	'!        dans "repertoire" ou le r�pertoire par d�faut (Serveur_Repertoire)!
	'!        les fichiers transmis.                                            !
	'!                                                                          !
	'! Parametres : POST 			 = Request.BinaryRead(Request.TotalBytes)         !
	'!						  repertoire = repertoire virtuel sur le serveur ou ""        !
	'!                                                                          !
	'! Valeur retournee : true si au moins un fichier est �crit, false sinon.   !
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
	
	'+--------------------------------------{ M�thode : repertoireServeur }-----+
	'!                                                                          !
	'! repertoireServeur(repertoire)                                            !
	'!                                                                          !
	'! role : Renseigne le param�tre "repServ" de l'objet donn�e membre "param".!
	'!                                                                          !
	'! Parametres : repertoire = chemin du r�pertoire sous la forme \rep\       !
	'!                                                                          !
	'! Valeur retournee :                                                       !
	'!                                                                          !
	'+--------------------------------------------------------------------------+
	function repertoireServeur(repertoire)
		param.repServ = repertoire
	end function
	
	'+---------------------------------{ M�thode : tailleFichiersUploades }-----+
	'!                                                                          !
	'! tailleFichiersUploades(taille)                                           !
	'!                                                                          !
	'! role : Renseigne le param�tre "tailleFichier" de l'objet donn�e membre   !
	'!        "param".                                                          !
	'!                                                                          !
	'! Parametres : taille = valeur num�rique en ko                             !
	'!                                                                          !
	'! Valeur retournee :                                                       !
	'!                                                                          !
	'+--------------------------------------------------------------------------+
	function tailleFichiersUploades(taille)
		param.tailleFichier = taille			
	end function
	
	'+-------------------------------------{ M�thode : extensionsUploadee }-----+
	'!                                                                          !
	'! extensionsUploadee(ext)                                                  !
	'!                                                                          !
	'! role : Renseigne le param�tre "extensions" de l'objet donn�e membre      !
	'!        "param".                                                          !
	'!                                                                          !
	'! Parametres : ext = chaine de caract�res listant les extensions, s�par�es !
	'!                    par des blancs. exemple : "txt htm html"              !
	'!                                                                          !
	'! Valeur retournee :                                                       !
	'!                                                                          !
	'+--------------------------------------------------------------------------+
	function extensionsUploadee(ext)
		param.extensions = ext		
	end function

	'+----------------------------------{ M�thode : extensionsNonUploadee }-----+
	'!                                                                          !
	'! extensionsNonUploadee(ext)                                               !
	'!                                                                          !
	'! role : Renseigne le param�tre "extToutSauf" de l'objet donn�e membre     !
	'!        "param".                                                          !
	'!                                                                          !
	'! Parametres : ext = chaine de caract�res listant les extensions, s�par�es !
	'!                    par des blancs. exemple : "txt htm html"              !
	'!                                                                          !
	'! Valeur retournee :                                                       !
	'!                                                                          !
	'+--------------------------------------------------------------------------+
	function extensionsNonUploadee(ext)
		param.extToutSauf = ext			
	end function

	'+------------------------------------------------{ M�thode : estNull }-----+
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

	'+-----------------------------------------------{ M�thode : cl_split }-----+
	'!                                                                          !
	'! cl_split(chaineBinaire,separateur)                                       !
	'!                                                                          !
	'! role : r�alise la commande split sur une chaine binaire.                 !
	'!                                                                          !
	'! Parametres : chaineBinaire = chaine binaire � traiter.                   !
	'!              separateur    = chaine binaire � prendre en compte comme    !
	'!                              s�parateur.                                 !
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

	'+-------------------------------------{ M�thode : NbreFichiersEcrits }-----+
	'!                                                                          !
	'! NbreFichiersEcrits()                                                     !
	'!                                                                          !
	'! role : Retourne le nombre de fichiers �crits.                            !
	'!                                                                          !
	'! Parametres :                                                             !
	'!                                                                          !
	'! Valeur retournee : nombre de fichiers �crits.                            !
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

	'+--------------------------------------{ M�thode : NbreTotalFichiers }-----+
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

	'+--------------------------------------{ M�thode : fichiersNonEcrits }-----+
	'!                                                                          !
	'! fichiersNonEcrits()                                                      !
	'!                                                                          !
	'! role : Retourne un tableau contenant tous les fichiers non upload�s.     !
	'!                                                                          !
	'! Parametres :                                                             !
	'!                                                                          !
	'! Valeur retournee : tableau contenant les fichier non upload�s.           !
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

	'+-----------------------------------------{ M�thode : fichiersEcrits }-----+
	'!                                                                          !
	'! fichiersEcrits()                                                         !
	'!                                                                          !
	'! role : Retourne un tableau contenant tous les fichiers upload�s.         !
	'!                                                                          !
	'! Parametres :                                                             !
	'!                                                                          !
	'! Valeur retournee : tableau contenant les fichiers upload�s.              !
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
		Response.write "*****************************<br>"
'		Response.write "TotalBytes 				: " & binaireTOascii(TotalBytes) & "<br>"
		param.AfficheObjet()
		fichier.AfficheObjet()


	end function

end class	
%>