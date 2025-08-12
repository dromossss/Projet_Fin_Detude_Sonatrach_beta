Explication des programmes
Ce dépôt contient plusieurs scripts Python pour manipuler des fichiers Excel via une interface Tkinter.

Description des programmes
PROG01
Crée une fenêtre Tkinter pour permettre à l'utilisateur de choisir deux fichiers Excel à combiner.

Lit les fichiers choisis dans des DataFrames distincts et les combine verticalement.

Analyse le fichier combiné et remplace les cellules vides par la valeur "VIDE".

Enregistre le fichier combiné sous le nom combined_file.xlsx.

PROG02
Utilise Tkinter pour créer une fenêtre demandant à l'utilisateur la référence d'une cellule (format : lettre + numéro, ex. A1) où il souhaite ajouter une donnée.

Ajoute la donnée introduite dans la cellule choisie.

Sauvegarde le fichier modifié.
Note : la fenêtre s'ouvre mais l'ajout de la valeur doit encore être vérifié.

PROG03
Lit deux cellules dans les fichiers choisis.

Crée deux nouvelles cellules :

Une contenant le texte "Total Global" (par exemple en cellule E9).

Une autre affichant le montant total (par exemple en cellule F9) au format monétaire.

Sauvegarde le fichier modifié.

PROGDOUBLONS
Crée une fenêtre Tkinter demandant à l'utilisateur de sélectionner un fichier Excel.

Demande à l'utilisateur de choisir la colonne à analyser.

Analyse la colonne pour détecter les doublons et les met en évidence (avec surlignage).

Sauvegarde le fichier modifié avec les doublons marqués.

Utilisation
Le programme principal fait appel à ces différentes fonctions pour appliquer une succession de traitements sur les fichiers Excel.

Notes additionnelles
Il est prévu d’ajouter une liaison avec une base de données (PHP/HTML ou SQL via PDO).

L’intégration future vise à automatiser les traitements via une base de données connectée.
