<%
'+--------------------------------------------{ Objet : ctrl_filename }-----+
'!                                                                          !
'! role : Cet objet gère la copie de fichiers sur le serveur.               !
'!                                                                          !
'! Méthodes publics :                                                       !
'! ------------------                                                       !
'!  construire(chaine)	 	: constructeur ou chaine est une occurence du     !
'!                          controle filename.                              !
'!  estNull()							: true si l'objet est null                        !
'!                        	liste d'objets erreurs.                         !
'!	ajoutFilename(tete_chaine)                                              !
'!												: chaine l'objet à la liste dont la tete est      !
'!                          "tete_chaine".                                  !
'!  estNull()							: true si l'objet est null                        !
'!                        	liste d'objets erreurs.                         !
'! 	estEcrit()						: true si le fichier a été uploadé.               !
'!  EcrireSurServeur(repertoire)                                            !
'!                        : écrit le fichier sur le serveur à l'endroit     !
'!                          indiqué par la constante Serveur_Repertoire du  !
'!                          fichier "upload.inc" si le paramètre repertoire !
'!                          est null.                                       !
'!  AfficheObjet()        : affiche les données membres de tous les objets  !
'!                          chainés à partir du courant sauf contenu.       !
'!						                                                              !
'! Méthodes privées :                                                       !
'! ------------------                                                       !
'!	sub class_initialize	: constructeur par défaut. Initialise à null les  !
'!                          données                                         !
'! 	sub class_terminate		: détruit tous les objets chainés                 !
'!	recherchePosition(position_debut,chaine1,chaine2)                       !
'!                        : retourne la position juste après "chaine1" dans !
'!                          "chaine2". La recherche debute à                !
'!                          "position-chaine"                               !
'!  ChargeExtension()	    : renseigne la donnée membre extension.           !
'!  ChargeNomRepExt(chaine)                                                 !
'!									      : renseigne les données membres nom, repertoire et!
'!                          extension.                                      !
'!  ChargeContenu(chaine) : renseigne la donnée membre contenu.             !
'!  controle(param)       : controle les éléments du fichier avant l'upload.!
'!						                                                              !
'+--------------------------------------------------------------------------+
class ctrl_filename
	'// Données membres
	public content_disposition
	public content_type
	public nom
	public extension
	public repertoire
	public contenu
	public temoin_ecrit  					'1 = fichier écrit, 0 = fichier non écrit
	public suivant

	'// Méthodes
	'+---------------------------------------{ Méthode : class_initialize }-----+
	'!                                                                          !
	'! class_initialize                                                         !
	'!                                                                          !
	'! role : Constructeur d'un l'objet ctrl_filename à null.                   !
	'!                                                                          !
	'! Parametres :                                                             !
	'!                                                                          !
	'! Valeur retournee :                                                       !
	'!                                                                          !
	'+--------------------------------------------------------------------------+
	private sub class_initialize
		content_disposition	= ""
		content_type				= ""
		nom									= ""
		extension						= ""
		repertoire					= ""
		contenu							= ""
		temoin_ecrit				= 0
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
		if isObject(suivant) then
			set suivant = nothing
		end if
	end sub
	
	'+---------------------------------------------{ Méthode : construire }-----+
	'!                                                                          !
	'! construire(chaine)                                                       !
	'!                                                                          !
	'! role : Constructeur d'un l'objet avec données transmises par un contrôle !
	'!        Filename. Ce constructeur est appelé pour chaque élément du       !
	'!        controle filename.                                                !
	'!                                                                          !
	'! Parametres : chaine = occurence du controle filename. (entre séparateur) !
	'!                                                                          !
	'! Valeur retournee : true si un fichier a été copié, false sinon.          !
	'!                                                                          !
	'+--------------------------------------------------------------------------+
	function construire(chaine)
		set suivant = new ctrl_filename

		if ChargeNomRepExt(chaine) then
			ChargeContenu(chaine)
			construire=true
			
		else
			construire=false
		end if
				
	end function

	'+------------------------------------------{ Méthode : ajoutFilename }-----+
	'!                                                                          !
	'! ajoutFilename(tete_chaine)                                               !
	'!                                                                          !
	'! role : Ajoute un objet de meme classe à la liste des objets chainés.     !
	'!                                                                          !
	'! Parametres :  tete_chaine		= tete de chaine                            !
	'!                                                                          !
	'! Valeur retournee : tete de chaine                                        !
	'!                                                                          !
	'+--------------------------------------------------------------------------+
	function ajoutFilename(tete_chaine)

		' gestion du chainage des objets 
		if not me.estNull then
		
			if not tete_chaine.estNull() then
				set filename_en_cours = tete_chaine
				while not filename_en_cours.suivant.estNull()
					set filename_en_cours=filename_en_cours.suivant
				wend
		
				set filename_en_cours.suivant=me
			else
				set tete_chaine=me
			end if
			
		end if
		set ajoutFilename=tete_chaine

	end function

	'+------------------------------------------------{ Méthode : estNull }-----+
	'!                                                                          !
	'! estNull()                                                                !
	'!                                                                          !
	'! role : Test si l'objet erreur est un objet null (contruit avec le        !
	'!        constructeur construire_null)                                     !
	'!                                                                          !
	'! Parametres :                                                             !
	'!                                                                          !
	'! Valeur retournee : true	: l'objet est null                              !
	'!                    false	: l'objet n'est pas null                        !
	'!                                                                          !
	'+--------------------------------------------------------------------------+
	function estNull()

		if	nom	= "" then
				estNull=true
		else
				estNull=false
		end if
	end function

	'+-----------------------------------------------{ Méthode : estEcrit }-----+
	'!                                                                          !
	'! estEcrit()                                                               !
	'!                                                                          !
	'! role : Retourne true si l'objet courant a été uploadé.                   !
	'!                                                                          !
	'! Parametres :                                                             !
	'!                                                                          !
	'! Valeur retournee : true	: le fichier a été uploadé                      !
	'!                                                                          !
	'+--------------------------------------------------------------------------+
	function estEcrit()
		if temoin_ecrit = 1 then
			estEcrit = true
		else
			estEcrit = false
		end if

	end function

	'+--------------------------------------{ Méthode : recherchePosition }-----+
	'!                                                                          !
	'! recherchePosition(position_debut,chaine,chaine_recherchee)               !
	'!                                                                          !
	'! role : retourne la position juste après "chaine_recherchee" dans         !
	'!        "chaine". La recherche debute à "position_debut".                 !
	'!                                                                          !
	'! Parametres : position_debut    = position à partir de laquelle il faut   !
	'!                                  rechercher "chaine".                    !
	'!              chaine				    = chaine dans laquelle la recherche est   !
	'!                                  effectuée.                              !
	'!              chaine_recherchee = chaine à rechercher.                    !
	'!                                                                          !
	'! Valeur retournee : position ou "0" si la chaine n est pas trouvée.       !
	'!                                                                          !
	'+--------------------------------------------------------------------------+
	private function recherchePosition(position_debut,chaine,chaine_recherchee)
	
	  PosDebutFic = Instrb(position_debut, chaine, chaine_recherchee)

		' -------------------
		' On lui ajoute ensuite la longueur du terme filename=" ce qui nous permet d'avoir la position de début du nom du fichier (PosDebutFic)
 	  recherchePosition = PosDebutFic+Lenb(chaine_recherchee)

	end function

	'+----------------------------------------{ Méthode : ChargeExtension }-----+
	'!                                                                          !
	'! ChargeExtension()                                                        !
	'!                                                                          !
	'! role : Renseigne la donnée membre extension.                             !
	'!                                                                          !
	'! Parametres :                                                             !
	'!                                                                          !
	'! Valeur retournee :                                                       !
	'!                                                                          !
	'+--------------------------------------------------------------------------+
	private function ChargeExtension()
		if nom <> "" then
			tab=split(nom,".",-1,1)
			extension = tab(ubound(tab))	 
		else
			extension=""
		end if
	end function

	'+----------------------------------------{ Méthode : ChargeNomRepExt }-----+
	'!                                                                          !
	'! ChargeNomRepExt(chaine)                                                  !
	'!                                                                          !
	'! role : Renseigne les données membres nom, repertoire et extension.       !
	'!                                                                          !
	'! Parametres : chaine = occurence du filename concernée                    !
	'!                                                                          !
	'! Valeur retournee : true si les données sont bien renseignées.            !
	'!                                                                          !
	'+--------------------------------------------------------------------------+
	private function ChargeNomRepExt(chainebinaire)
	
		ChargeNomRepExt=false

		PosDebutFilename		= recherchePosition(1,chainebinaire,asciiTObinaire("filename=" & chr(34)))
		if PosDebutFilename <> 0 then
			
			PosDebutContentType	= recherchePosition(1,chainebinaire,asciiTObinaire("Content-Type:"))
			
	
			' On trouve la position de la fin du nom du fichier à partir de la position 
			' du début du terme Content-Type: à laquelle on retire trois octets 
			' (un espace, une " et la première lettre du terme)
			PosFinFilename = PosDebutContentType - 3  
	
			if PosFinFilename > PosDebutFilename then
				nom = binaireTOascii(midb(chainebinaire,PosDebutFilename,(PosFinFilename-PosDebutFilename))) 
				
				tab=split(nom,"\",-1,1)
				 
				if(ubound(tab) > 0) then 
					tab2=split(tab(ubound(tab)),chr(34),-1,1)
					nom=tab2(0)
					repertoire=""
			  	for i = 0 to ubound(tab)-1
						repertoire = repertoire&tab(i)&"\"
					next
				else
					nom 				= ""
					repertoire 	= ""
				end if
		
				ChargeExtension()
	
				ChargeNomRepExt=true
			end if
		end if

	end function
	
	'+------------------------------------------{ Méthode : ChargeContenu }-----+
	'!                                                                          !
	'! ChargeContenu(chaine)                                                    !
	'!                                                                          !
	'! role : Renseigne la donnée membre contenu.                               !
	'!                                                                          !
	'! Parametres : chaine = occurence du filename concernée                    !
	'!                                                                          !
	'! Valeur retournee :                                                       !
	'!                                                                          !
	'+--------------------------------------------------------------------------+
	private function ChargeContenu(chainebinaire)
	
		' -------------------   
		' On cherche la position de début du contenu du fichier en sautant les blancs
 
		PosDebutContentType	= recherchePosition(1,chainebinaire,asciiTObinaire("Content-Type:"))

   	PosFinContentType = Instrb(PosDebutContentType, chainebinaire, asciiTObinaire(chr(13)))

  	if PosFinContentType <> 0 then
   		PosDebutContenu = PosFinContentType + 4 
	  end if

		content_type = binaireTOascii(midb(chainebinaire, PosDebutContentType, PosFinContentType-PosDebutContentType))

   	contenu = midb(chainebinaire, PosDebutContenu, lenb(chainebinaire)) 
	end function

	
	'+-----------------------------------------------{ Méthode : controle }-----+
	'!                                                                          !
	'! controle(param)                                                          !
	'!                                                                          !
	'! role : Contrôle les éléments du fichier à uploader par rapport aux       !
	'!        paramètres de l'objet "parametre".                                !
	'!                                                                          !
	'! Parametres : param = objet de type parametre                             !
	'!                                                                          !
	'! Valeur retournee : true si le controle est valide, false sinon.          !
	'!                                                                          !
	'+--------------------------------------------------------------------------+
	private function controle(param)

		controle = true

		if not param.controleTaille(len(contenu)) then controle = false
		
		if not param.controleExt(extension) then controle = false
		
		if not param.controleExtSauf(extension) then controle = false

	end function

	'+---------------------------------------{ Méthode : EcrireSurServeur }-----+
	'!                                                                          !
	'! EcrireSurServeur(repertoire)                                             !
	'!                                                                          !
	'! role : Ecrit le fichier sur le serveur à l'endroit indiqué par la        !
	'!        constante Serveur_Repertoire du fichier "upload.inc" si le        !
	'!        paramètre repertoire est null.                                    !
	'!                                                                          !
	'! Parametres : repertoire = repertoire virtuel sur le serveur ou ""        !
	'!                                                                          !
	'! Valeur retournee :                                                       !
	'!                                                                          !
	'+--------------------------------------------------------------------------+
	function EcrireSurServeur(param)
  	EcrireSurServeur = 0

	  if not me.estNull() then 
			if lenb(contenu) <= (param.tailleLimite()*1000) then
	   		' NouveauFic = Server.MapPath("\") & param.repertoireServ & nom
	   		NouveauFic = user_uploadFolder & "\" & nom
				if controle(param) then
			   	Set FileObject = Server.CreateObject("Scripting.FileSystemObject")
'			   	if FileObject.FolderExists(Server.MapPath("\") & param.repertoireServ) then
			   	if FileObject.FolderExists(user_uploadFolder & "\") then
					   	Set Out=FileObject.CreateTextFile(NouveauFic, True)
				   	For I = 1 to Lenb(contenu)
				   		Out.Write chr(ascb(midb(contenu,I,1)))
				   	Next
				   	Out.close
				   	Set Out=nothing
						temoin_ecrit = 1
					end if
				end if
		  end if

			if not suivant.estNull() then suivant.EcrireSurServeur(param)
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
			Response.write "-------------------------------------<br>"
		if estEcrit() then
			Response.write "nom			 					: " & nom & "  -> est copié sur le serveur.<br>"
		else
			Response.write "nom			 					: " & nom & "  -> n'est pas copié sur le serveur.<br>"
		end if
		Response.write "extension 				: " & extension & "<br>"
		Response.write "repertoire 				: " & repertoire & "<br>"
		Response.write "content_type			: " & content_type & "<br>"
		Response.write "taille du fichier : " & lenb(contenu) & " octets<br>"

		if not suivant.estNull() then 
			suivant.afficheObjet()
		end if
	end function

end class	
%>