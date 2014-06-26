<%
'+--------------------------------------------{ Objet : ctrl_filename }-----+
'!                                                                          !
'! role : Cet objet g�re la copie de fichiers sur le serveur.               !
'!                                                                          !
'! M�thodes publics :                                                       !
'! ------------------                                                       !
'!  construire(chaine)	 	: constructeur ou chaine est une occurence du     !
'!                          controle filename.                              !
'!  estNull()							: true si l'objet est null                        !
'!                        	liste d'objets erreurs.                         !
'!	ajoutFilename(tete_chaine)                                              !
'!												: chaine l'objet � la liste dont la tete est      !
'!                          "tete_chaine".                                  !
'!  estNull()							: true si l'objet est null                        !
'!                        	liste d'objets erreurs.                         !
'! 	estEcrit()						: true si le fichier a �t� upload�.               !
'!  EcrireSurServeur(repertoire)                                            !
'!                        : �crit le fichier sur le serveur � l'endroit     !
'!                          indiqu� par la constante Serveur_Repertoire du  !
'!                          fichier "upload.inc" si le param�tre repertoire !
'!                          est null.                                       !
'!  AfficheObjet()        : affiche les donn�es membres de tous les objets  !
'!                          chain�s � partir du courant sauf contenu.       !
'!						                                                              !
'! M�thodes priv�es :                                                       !
'! ------------------                                                       !
'!	sub class_initialize	: constructeur par d�faut. Initialise � null les  !
'!                          donn�es                                         !
'! 	sub class_terminate		: d�truit tous les objets chain�s                 !
'!	recherchePosition(position_debut,chaine1,chaine2)                       !
'!                        : retourne la position juste apr�s "chaine1" dans !
'!                          "chaine2". La recherche debute �                !
'!                          "position-chaine"                               !
'!  ChargeExtension()	    : renseigne la donn�e membre extension.           !
'!  ChargeNomRepExt(chaine)                                                 !
'!									      : renseigne les donn�es membres nom, repertoire et!
'!                          extension.                                      !
'!  ChargeContenu(chaine) : renseigne la donn�e membre contenu.             !
'!  controle(param)       : controle les �l�ments du fichier avant l'upload.!
'!						                                                              !
'+--------------------------------------------------------------------------+
class ctrl_filename
	'// Donn�es membres
	public content_disposition
	public content_type
	public nom
	public extension
	public repertoire
	public contenu
	public temoin_ecrit  					'1 = fichier �crit, 0 = fichier non �crit
	public suivant

	'// M�thodes
	'+---------------------------------------{ M�thode : class_initialize }-----+
	'!                                                                          !
	'! class_initialize                                                         !
	'!                                                                          !
	'! role : Constructeur d'un l'objet ctrl_filename � null.                   !
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
		if isObject(suivant) then
			set suivant = nothing
		end if
	end sub
	
	'+---------------------------------------------{ M�thode : construire }-----+
	'!                                                                          !
	'! construire(chaine)                                                       !
	'!                                                                          !
	'! role : Constructeur d'un l'objet avec donn�es transmises par un contr�le !
	'!        Filename. Ce constructeur est appel� pour chaque �l�ment du       !
	'!        controle filename.                                                !
	'!                                                                          !
	'! Parametres : chaine = occurence du controle filename. (entre s�parateur) !
	'!                                                                          !
	'! Valeur retournee : true si un fichier a �t� copi�, false sinon.          !
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

	'+------------------------------------------{ M�thode : ajoutFilename }-----+
	'!                                                                          !
	'! ajoutFilename(tete_chaine)                                               !
	'!                                                                          !
	'! role : Ajoute un objet de meme classe � la liste des objets chain�s.     !
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

	'+------------------------------------------------{ M�thode : estNull }-----+
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

	'+-----------------------------------------------{ M�thode : estEcrit }-----+
	'!                                                                          !
	'! estEcrit()                                                               !
	'!                                                                          !
	'! role : Retourne true si l'objet courant a �t� upload�.                   !
	'!                                                                          !
	'! Parametres :                                                             !
	'!                                                                          !
	'! Valeur retournee : true	: le fichier a �t� upload�                      !
	'!                                                                          !
	'+--------------------------------------------------------------------------+
	function estEcrit()
		if temoin_ecrit = 1 then
			estEcrit = true
		else
			estEcrit = false
		end if

	end function

	'+--------------------------------------{ M�thode : recherchePosition }-----+
	'!                                                                          !
	'! recherchePosition(position_debut,chaine,chaine_recherchee)               !
	'!                                                                          !
	'! role : retourne la position juste apr�s "chaine_recherchee" dans         !
	'!        "chaine". La recherche debute � "position_debut".                 !
	'!                                                                          !
	'! Parametres : position_debut    = position � partir de laquelle il faut   !
	'!                                  rechercher "chaine".                    !
	'!              chaine				    = chaine dans laquelle la recherche est   !
	'!                                  effectu�e.                              !
	'!              chaine_recherchee = chaine � rechercher.                    !
	'!                                                                          !
	'! Valeur retournee : position ou "0" si la chaine n est pas trouv�e.       !
	'!                                                                          !
	'+--------------------------------------------------------------------------+
	private function recherchePosition(position_debut,chaine,chaine_recherchee)
	
	  PosDebutFic = Instrb(position_debut, chaine, chaine_recherchee)

		' -------------------
		' On lui ajoute ensuite la longueur du terme filename=" ce qui nous permet d'avoir la position de d�but du nom du fichier (PosDebutFic)
 	  recherchePosition = PosDebutFic+Lenb(chaine_recherchee)

	end function

	'+----------------------------------------{ M�thode : ChargeExtension }-----+
	'!                                                                          !
	'! ChargeExtension()                                                        !
	'!                                                                          !
	'! role : Renseigne la donn�e membre extension.                             !
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

	'+----------------------------------------{ M�thode : ChargeNomRepExt }-----+
	'!                                                                          !
	'! ChargeNomRepExt(chaine)                                                  !
	'!                                                                          !
	'! role : Renseigne les donn�es membres nom, repertoire et extension.       !
	'!                                                                          !
	'! Parametres : chaine = occurence du filename concern�e                    !
	'!                                                                          !
	'! Valeur retournee : true si les donn�es sont bien renseign�es.            !
	'!                                                                          !
	'+--------------------------------------------------------------------------+
	private function ChargeNomRepExt(chainebinaire)
	
		ChargeNomRepExt=false

		PosDebutFilename		= recherchePosition(1,chainebinaire,asciiTObinaire("filename=" & chr(34)))
		if PosDebutFilename <> 0 then
			
			PosDebutContentType	= recherchePosition(1,chainebinaire,asciiTObinaire("Content-Type:"))
			
	
			' On trouve la position de la fin du nom du fichier � partir de la position 
			' du d�but du terme Content-Type: � laquelle on retire trois octets 
			' (un espace, une " et la premi�re lettre du terme)
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
	
	'+------------------------------------------{ M�thode : ChargeContenu }-----+
	'!                                                                          !
	'! ChargeContenu(chaine)                                                    !
	'!                                                                          !
	'! role : Renseigne la donn�e membre contenu.                               !
	'!                                                                          !
	'! Parametres : chaine = occurence du filename concern�e                    !
	'!                                                                          !
	'! Valeur retournee :                                                       !
	'!                                                                          !
	'+--------------------------------------------------------------------------+
	private function ChargeContenu(chainebinaire)
	
		' -------------------   
		' On cherche la position de d�but du contenu du fichier en sautant les blancs
 
		PosDebutContentType	= recherchePosition(1,chainebinaire,asciiTObinaire("Content-Type:"))

   	PosFinContentType = Instrb(PosDebutContentType, chainebinaire, asciiTObinaire(chr(13)))

  	if PosFinContentType <> 0 then
   		PosDebutContenu = PosFinContentType + 4 
	  end if

		content_type = binaireTOascii(midb(chainebinaire, PosDebutContentType, PosFinContentType-PosDebutContentType))

   	contenu = midb(chainebinaire, PosDebutContenu, lenb(chainebinaire)) 
	end function

	
	'+-----------------------------------------------{ M�thode : controle }-----+
	'!                                                                          !
	'! controle(param)                                                          !
	'!                                                                          !
	'! role : Contr�le les �l�ments du fichier � uploader par rapport aux       !
	'!        param�tres de l'objet "parametre".                                !
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

	'+---------------------------------------{ M�thode : EcrireSurServeur }-----+
	'!                                                                          !
	'! EcrireSurServeur(repertoire)                                             !
	'!                                                                          !
	'! role : Ecrit le fichier sur le serveur � l'endroit indiqu� par la        !
	'!        constante Serveur_Repertoire du fichier "upload.inc" si le        !
	'!        param�tre repertoire est null.                                    !
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
			Response.write "-------------------------------------<br>"
		if estEcrit() then
			Response.write "nom			 					: " & nom & "  -> est copi� sur le serveur.<br>"
		else
			Response.write "nom			 					: " & nom & "  -> n'est pas copi� sur le serveur.<br>"
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