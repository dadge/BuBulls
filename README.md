
# BuBulls

## Pr�requis
T�l�chargez l'outil la version derni�re version compil�e de l'outil [ici](https://github.com/semsitivity/BuBulls/releases).
L�outil ne fonctionne que sous windows 10 � jour.
Pour le faire fonctionner, il faut bien s�r le r�pertoire contenant le programme BuBulls.exe (ainsi que les fichiers .dll).
Il vous faudra Excel et Word (version sup�rieure � office 2003 SP3, mais le plus r�cent sera le mieux).
Il faut aussi le fichier de mod�le de base : BuBulls_template.docx
Et il faut encore le fichier Excel pour l�encodage : Modele_Bulletins.xlsx

## Principe
Le principe du programme est d�analyser le fichier Excel et d�en retirer la liste des �l�ves, des mati�res et de leurs comp�tences et �videment les r�sultats des �l�ves.
Pour l�encodage des r�sultats dans l�Excel :
1.	La liste des �l�ves doit �tre entr�e dans l�onglet El�ves.
2.	La liste des comp�tences doit �tre entr�e dans l�onglet Mati�res. 
La premi�re colonne correspond aux mati�res et la seconde aux comp�tences associ�es � la mati�re de la premi�re colonne. Il n�y a aucun probl�me � rajouter ou supprimer des lignes, par contre cela peut d�caler les encodages qui ont �t� fat dans l�onglet acquisition.
La seule contrainte est que le nom de la mati�re doit correspondre exactement (� la majuscule ou accent pr�t) aux noms des mati�res entour�s de {{ }} dans le mod�le de base.
3.	L�onglet acquisition se construit automatiquement sur base de la liste d��l�ves et la liste des comp�tences. Le tableau doit �tre rempli de 0 (non acquis), 1 (acquis) et 2 (Acquis avec transfert).
Un fois l�encodage de l�Excel termin�, il est alors possible d�utiliser le programme.
Dans un premier temps, le programme va utiliser le mod�le de base Modele_Bulletins.xlsx et va remplir les tableaux de mati�res par les comp�tences qui y sont associ�es dans l�Excel.  La mise en page du mod�le de base sera donc conserv�e, ainsi il est possible d�apporter des modifications dans le mod�le de base. Une fois encore, il est important que chaque nom de mati�re concorde exactement avec les mati�res list�es dans l�Excel et ce afin d�y ins�rer les comp�tences correspondantes.
Pour information, les 2 textes {{fn}} et {{ln}} seront respectivement remplac�s par le pr�nom et nom de l��l�ve. Les {{mati�res}} seront remplac� par les listes de comp�tences, et les {{X}} par les r�sultats encod�s. Il est donc essentiel de ne pas les supprimer du mod�le.

Dans un premier temps, le programme va g�n�rer un mod�le word interm�diaire. Ce mod�le interm�diaire peut �tre vu comme un bulletin anonyme, c�est-�-dire qu�il doit ressembler au plus possible � la mise en page finale du bulletin. Il s�agit du mod�le de base peupl� des comp�tences. Durant cette �tape, il est important de finaliser la mise en page et notamment la bonne mise en forme des sauts de pages et s�paration des mati�res par pages.
Un fois cette �tape termin�e et la mise en page d�finitive choisie (contenant les comp�tences), le programme peut g�n�rer tous les bulletins conform�ment � ce qui a �t� saisi dans l�Excel.

## Utilisation du programme
D�marrez le programme en double cliquant sur BuBulls.exe.
Veillez � ce qu�aucun fichier Word ou Excel ne soit ouvert, en tout cas surtout pas le mod�le de base ni l�Excel d�encodage.
Le programme vous invite � faire un drag&drop (glisser d�placer) du fichier Word du mod�le de base jusque dans la case intitul�e � 1. Dropper le template de base ici �.
Une fois fait, le programme vous invite � faire un drag&drop du fichier d�encodage Excel dans la case � 2.Dropper l�Excel ici �. Cela va entrainer l�analyse des bulletins qui peut prendre quelques secondes.
D�s que l�analyse est termin�e, et donc que le mod�le interm�diaire est g�n�r�, le programme va ouvrir le mod�le interm�diaire dans Word pour vous. 
Editez le document et adaptez la mise en page et les sauts de pages � votre convenance maintenant que le document contient toutes les comp�tences. V�rifiez rapidement aussi que toutes les mati�res sont pr�sentes. Une fois ce travail termin�, sauvegardez et fermez Word.
A la fermeture, le programme va vous proposer un bouton pour g�n�rer les bulletins. Il suffit de cliquer dessus et les bulletins vont �tre g�n�r�s dans le m�me r�pertoire d�o� vient l�Excel d�encodage dans un sous r�pertoire � bulletins �. Le r�pertoire s�ouvrira automatiquement � la fin de la g�n�ration des bulletins.
