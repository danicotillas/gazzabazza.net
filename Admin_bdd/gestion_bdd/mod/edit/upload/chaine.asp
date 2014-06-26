<%

' prototype des fonctions utiles
' ------------------------------
'function ControleChaineNonVide(chaine)
'function Trace(chaine)
'function binaireTOascii(chaineBinaire)
'function asciiTObinaire(chaine)

'+--------------------------------------------{ ControleChaineNonVide }-----+
'!                                                                          !
'! ControleChaineNonVide(chaine)                                            !
'!                                                                          !
'! role : Controle que la chaine passée en paramètre n'est pas vide. Si la  !
'!        chaine est vide, un nouvelle objet erreur est chaines aux autres. !
'!                                                                          !
'! Parametres : chaine 			= chaine de caractères à controler              !
'!              listeErreur 	= pointeur sur le premier element de la liste !
'!                             des erreurs.                                 !
'!              ajoutDsListeErreur = 1 pour générer une nouvelle erreur si  !
'!                                     la chaine est vide.                  !
'!                                   0 pou                                  !
'!                                                                          !
'! Valeur retournee : pointeur sur le premier element de la liste chainee   !
'!                    des objet erreurs ou "null" s'il n'y a pas d'erreur.  !
'!                                                                          !
'+--------------------------------------------------------------------------+
function ControleChaineNonVide(nomChamp,chaine,premierErreurs) 

	dim Erreur
	set Erreur = new Erreurs

	if len(chaine) < 1 then
		Erreur.construire nomChamp,chaine,Champ_Obligatoire,null
		set premierErreurs = Erreur.ajoutErreur(premierErreurs)
	end if
	
	set ControleChaineNonVide=premierErreurs

end function

'+------------------------------------------------------------{ trace }-----+
'!                                                                          !
'! trace(chaine)                                                            !
'!                                                                          !
'! role : Affiche chaine suivi de la balise <br>.                           !
'!                                                                          !
'! Parametres : chaine = chaine de caractères à afficher                    !
'!                                                                          !
'! Valeur retournee :                                                       !
'!                                                                          !
'+--------------------------------------------------------------------------+
function trace(chaine) 

	if chaine <> "" then
		Response.Write chaine & "<br>"
	else
		Response.Write "<br>"
	end if
end function

'+-----------------------------------------{ Méthode : binaireTOascii }-----+
'!                                                                          !
'! binaireTOascii(chaineBinaire)                                            !
'!                                                                          !
'! role : Transforme une chaine binaire en chaine ASCII.                    !
'!                                                                          !
'! Parametres : chaineBinaire = chaine binaire à transformer.               !
'!                                                                          !
'! Valeur retournee : chaine ASCII                                          !
'!                                                                          !
'+--------------------------------------------------------------------------+
function binaireTOascii(chaineBinaire)
	ContenuAscii = ""
	for Z = 1 to LenB(chaineBinaire)
		ContenuAscii = ContenuAscii & chr(ASCB(MidB(chaineBinaire, Z, 1)))
	next
	binaireTOascii = ContenuAscii
end function

'+-----------------------------------------{ Méthode : asciiTObinaire }-----+
'!                                                                          !
'! asciiTObinaire(chaine)                                                   !
'!                                                                          !
'! role : Transforme une chaine binaire en chaine ASCII.                    !
'!                                                                          !
'! Parametres : chaineBinaire = chaine binaire à transformer.               !
'!                                                                          !
'! Valeur retournee : chaine ASCII                                          !
'!                                                                          !
'+--------------------------------------------------------------------------+
function asciiTObinaire(chaine)
  For I=1 to len(chaine)
    B = B & ChrB(Asc(Mid(chaine,I,1)))
  Next
  asciiTObinaire = B
End Function
	
%>

