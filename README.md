
# BuBulls

## Prérequis
Téléchargez l'outil la version dernière version compilée de l'outil [ici](https://github.com/semsitivity/BuBulls/releases).
L’outil ne fonctionne que sous windows 10 à jour.
Pour le faire fonctionner, il faut bien sûr le répertoire contenant le programme BuBulls.exe (ainsi que les fichiers .dll).
Il vous faudra Excel et Word (version supérieure à office 2003 SP3, mais le plus récent sera le mieux).
Il faut aussi le fichier de modèle de base : BuBulls_template.docx
Et il faut encore le fichier Excel pour l’encodage : Modele_Bulletins.xlsx

## Principe
Le principe du programme est d’analyser le fichier Excel et d’en retirer la liste des élèves, des matières et de leurs compétences et évidement les résultats des élèves.
Pour l’encodage des résultats dans l’Excel :
1.	La liste des élèves doit être entrée dans l’onglet Elèves.
2.	La liste des compétences doit être entrée dans l’onglet Matières. 
La première colonne correspond aux matières et la seconde aux compétences associées à la matière de la première colonne. Il n’y a aucun problème à rajouter ou supprimer des lignes, par contre cela peut décaler les encodages qui ont été fat dans l’onglet acquisition.
La seule contrainte est que le nom de la matière doit correspondre exactement (à la majuscule ou accent prêt) aux noms des matières entourés de {{ }} dans le modèle de base.
3.	L’onglet acquisition se construit automatiquement sur base de la liste d’élèves et la liste des compétences. Le tableau doit être rempli de 0 (non acquis), 1 (acquis) et 2 (Acquis avec transfert).
Un fois l’encodage de l’Excel terminé, il est alors possible d’utiliser le programme.
Dans un premier temps, le programme va utiliser le modèle de base Modele_Bulletins.xlsx et va remplir les tableaux de matières par les compétences qui y sont associées dans l’Excel.  La mise en page du modèle de base sera donc conservée, ainsi il est possible d’apporter des modifications dans le modèle de base. Une fois encore, il est important que chaque nom de matière concorde exactement avec les matières listées dans l’Excel et ce afin d’y insérer les compétences correspondantes.
Pour information, les 2 textes {{fn}} et {{ln}} seront respectivement remplacés par le prénom et nom de l’élève. Les {{matières}} seront remplacé par les listes de compétences, et les {{X}} par les résultats encodés. Il est donc essentiel de ne pas les supprimer du modèle.

Dans un premier temps, le programme va générer un modèle word intermédiaire. Ce modèle intermédiaire peut être vu comme un bulletin anonyme, c’est-à-dire qu’il doit ressembler au plus possible à la mise en page finale du bulletin. Il s’agit du modèle de base peuplé des compétences. Durant cette étape, il est important de finaliser la mise en page et notamment la bonne mise en forme des sauts de pages et séparation des matières par pages.
Un fois cette étape terminée et la mise en page définitive choisie (contenant les compétences), le programme peut générer tous les bulletins conformément à ce qui a été saisi dans l’Excel.

## Utilisation du programme
Démarrez le programme en double cliquant sur BuBulls.exe.
Veillez à ce qu’aucun fichier Word ou Excel ne soit ouvert, en tout cas surtout pas le modèle de base ni l’Excel d’encodage.
Le programme vous invite à faire un drag&drop (glisser déplacer) du fichier Word du modèle de base jusque dans la case intitulée « 1. Dropper le template de base ici ».
Une fois fait, le programme vous invite à faire un drag&drop du fichier d’encodage Excel dans la case « 2.Dropper l’Excel ici ». Cela va entrainer l’analyse des bulletins qui peut prendre quelques secondes.
Dès que l’analyse est terminée, et donc que le modèle intermédiaire est généré, le programme va ouvrir le modèle intermédiaire dans Word pour vous. 
Editez le document et adaptez la mise en page et les sauts de pages à votre convenance maintenant que le document contient toutes les compétences. Vérifiez rapidement aussi que toutes les matières sont présentes. Une fois ce travail terminé, sauvegardez et fermez Word.
A la fermeture, le programme va vous proposer un bouton pour générer les bulletins. Il suffit de cliquer dessus et les bulletins vont être générés dans le même répertoire d’où vient l’Excel d’encodage dans un sous répertoire « bulletins ». Le répertoire s’ouvrira automatiquement à la fin de la génération des bulletins.
