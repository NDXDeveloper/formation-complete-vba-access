🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 20.2. Protection par mot de passe de la base de données

La protection par mot de passe est la première ligne de défense d'une base `.accdb`, et la plus simple à mettre en œuvre. Comme nous l'avons vu en [section 20.1](01-modele-securite-access.md), il ne s'agit pas d'un système de comptes individuels : c'est un mot de passe unique, partagé par tous, qui commande l'ouverture du fichier. Sa fonction est double. D'une part il interdit l'ouverture de la base à quiconque ne le connaît pas. D'autre part, et c'est l'essentiel, il déclenche le chiffrement de l'intégralité du fichier sur le disque, rendant illisible le contenu brut des tables même pour quelqu'un qui ouvrirait le `.accdb` avec un éditeur hexadécimal ou un autre outil.

Cette section décrit comment poser, modifier et retirer ce mot de passe, à la fois par l'interface et par code, puis détaille les précautions indispensables.

## Une condition incontournable : l'ouverture en mode exclusif

Avant toute chose, retenez ceci : on ne peut définir, modifier ou supprimer un mot de passe que si la base est ouverte en mode exclusif. Tant que la base est ouverte en mode partagé (le mode par défaut), la commande de chiffrement reste grisée et inaccessible.

Pour ouvrir une base en exclusif via l'interface, on utilise la boîte de dialogue d'ouverture de fichier : après avoir sélectionné le fichier, on clique sur la petite flèche située à droite du bouton Ouvrir, puis on choisit Ouvrir en exclusif. Cette contrainte est logique : chiffrer ou déchiffrer un fichier suppose de le réécrire entièrement, ce qui est impossible si d'autres utilisateurs y sont connectés.

## Définir un mot de passe via l'interface

Une fois la base ouverte en exclusif, la marche à suivre est la suivante. On se rend dans l'onglet Fichier, puis dans la rubrique Informations, et l'on clique sur Chiffrer avec un mot de passe. Access demande alors de saisir le mot de passe deux fois pour confirmation, puis applique le chiffrement à l'ensemble du fichier. À la prochaine ouverture, le mot de passe sera exigé.

## Supprimer ou changer le mot de passe via l'interface

Le retrait suit la même logique. La base étant ouverte en exclusif (ce qui aura nécessité de fournir le mot de passe à l'ouverture), on retourne dans Fichier puis Informations, où la commande s'intitule désormais Déchiffrer la base de données. Access demande le mot de passe actuel, puis supprime le chiffrement.

Il n'existe pas de commande unique pour « changer » le mot de passe via l'interface : on procède en deux temps, en déchiffrant d'abord la base avec l'ancien mot de passe, puis en la rechiffrant avec le nouveau.

## Gérer le mot de passe par code (DAO)

La bibliothèque DAO permet d'automatiser ces opérations grâce à la méthode `NewPassword` de l'objet `Database`. Cette méthode prend deux arguments : l'ancien mot de passe et le nouveau. La convention est simple et élégante : une chaîne vide représente l'absence de mot de passe.

Là encore, la base concernée doit être ouverte en exclusif, ce qui correspond au deuxième argument de `OpenDatabase` positionné à `True`. On ne peut donc pas appliquer cette méthode à la base courante (`CurrentDb`), qui est ouverte en mode partagé : il faut ouvrir explicitement le fichier cible, qui doit être fermé par ailleurs.

Pour **poser** un mot de passe sur une base qui n'en a pas :

```vba
Dim db As DAO.Database
' 2e argument True = ouverture exclusive (obligatoire pour cette opération)
Set db = DBEngine.OpenDatabase("C:\Data\Donnees.accdb", True, False)

db.NewPassword "", "MonMotDePasseFort!2026"   ' "" = aucun mot de passe actuel

db.Close
Set db = Nothing
```

Pour **changer** un mot de passe existant, il faut d'abord ouvrir la base en fournissant le mot de passe actuel dans la chaîne de connexion (`;PWD=`), puis appeler `NewPassword` avec l'ancien et le nouveau :

```vba
Dim db As DAO.Database
Set db = DBEngine.OpenDatabase("C:\Data\Donnees.accdb", True, False, _
                               ";PWD=MonMotDePasseFort!2026")

db.NewPassword "MonMotDePasseFort!2026", "UnAutreMotDePasse#2027"

db.Close
Set db = Nothing
```

Pour **supprimer** un mot de passe, on passe une chaîne vide comme nouveau mot de passe :

```vba
Dim db As DAO.Database
Set db = DBEngine.OpenDatabase("C:\Data\Donnees.accdb", True, False, _
                               ";PWD=UnAutreMotDePasse#2027")

db.NewPassword "UnAutreMotDePasse#2027", ""   ' "" = plus de mot de passe

db.Close
Set db = Nothing
```

## Ouvrir une base protégée par code

Au quotidien, le besoin le plus fréquent n'est pas de manipuler le mot de passe mais simplement d'ouvrir une base protégée pour y travailler. La technique dépend de la bibliothèque utilisée.

En DAO, on renseigne le mot de passe dans le quatrième argument de `OpenDatabase`, cette fois en mode partagé (`False`) puisqu'il s'agit d'un accès normal et non d'une opération de chiffrement :

```vba
Dim db As DAO.Database
Set db = DBEngine.OpenDatabase("C:\Data\Donnees.accdb", _
                               False, _              ' partagé
                               False, _              ' lecture-écriture
                               ";PWD=MonMotDePasse")
' ... lecture / écriture sur db ...
db.Close
Set db = Nothing
```

En ADO, le mot de passe se place dans la chaîne de connexion via la propriété `Jet OLEDB:Database Password` :

```vba
Dim cn As ADODB.Connection
Set cn = New ADODB.Connection
cn.Open "Provider=Microsoft.ACE.OLEDB.12.0;" & _
        "Data Source=C:\Data\Donnees.accdb;" & _
        "Jet OLEDB:Database Password=MonMotDePasse;"
' ... travail via cn ...
cn.Close
Set cn = Nothing
```

Selon la version du moteur installée, le fournisseur peut s'écrire `Microsoft.ACE.OLEDB.12.0` ou `Microsoft.ACE.OLEDB.16.0`. Les chaînes de connexion sont détaillées au [chapitre 10](../10-ado-access/02-chaines-connexion-access.md) et en [annexe D](../annexes/d-chaines-connexion-ado-odbc.md).

## Le piège des tables liées vers une base protégée

Dans une architecture séparée en front-end et back-end (voir [chapitre 15](../15-multi-utilisateurs/01-architecture-front-end-back-end.md)), on chiffre logiquement le back-end, qui contient les données. Se pose alors une difficulté : comment les tables liées du front-end accèdent-elles à un back-end protégé ?

Lorsqu'on crée une table liée vers une base protégée, Access propose d'enregistrer le mot de passe. Si l'on accepte, ce mot de passe est stocké dans la propriété `Connect` de la table liée, à l'intérieur du front-end. Le problème de sécurité est sérieux : ce mot de passe y figure de façon directement lisible. Quiconque parvient à inspecter les tables liées du front-end peut donc retrouver le mot de passe du back-end, ce qui réduit considérablement l'intérêt du chiffrement.

```vba
' La propriété Connect d'une table liée vers un back-end protégé
' EXPOSE le mot de passe en clair dans le front-end.
Debug.Print tdf.Connect
' affiche par exemple :
' ;DATABASE=\\Serveur\Partage\BackEnd.accdb;PWD=MonMotDePasse
```

La parade habituelle consiste à ne pas mémoriser le mot de passe dans les liaisons, mais à reconnecter les tables au démarrage de l'application par code, en construisant la chaîne de connexion à la volée et sans laisser le mot de passe en clair dans un module non protégé. Cette reliaison programmatique est traitée en [section 21.4](../21-deploiement-distribution/04-reliaison-automatique-tables.md). Cela protège la chaîne stockée, mais rappelle une limite générale d'Access : un mot de passe que l'application doit elle-même connaître pour fonctionner finit toujours par exister quelque part dans le front-end.

## Force du chiffrement et choix de la méthode

La robustesse de cette protection a beaucoup évolué. Les anciens fichiers `.mdb` n'offraient qu'un encodage trivial, que des utilitaires gratuits retiraient en quelques secondes. Le format `.accdb` introduit avec Access 2007 a apporté un véritable chiffrement, puis Access 2010 l'a nettement renforcé en s'appuyant sur l'API cryptographique de Windows et un algorithme moderne. Sur une version récente d'Access, le chiffrement par mot de passe d'un `.accdb` constitue donc une protection sérieuse — à une condition décisive : que le mot de passe lui-même soit fort. Un mot de passe court ou prévisible reste vulnérable à une attaque par dictionnaire ou par force brute, le chiffrement le plus solide ne compensant jamais un mot de passe faible.

Access propose par ailleurs un réglage de la méthode de chiffrement, accessible dans les options de l'application (selon la version, dans la rubrique des paramètres du client, à la section consacrée au chiffrement). Deux choix s'offrent à vous : la méthode par défaut, plus sûre, et une méthode héritée, présentée comme assurant une meilleure compatibilité et de meilleures performances. Ce second choix existe parce que le chiffrement fort peut, dans certains scénarios multi-utilisateurs, dégrader les performances d'accès concurrent. Il faut en avoir conscience, mais l'arbitrage doit en principe pencher vers la sécurité : la méthode héritée est plus faible et ne devrait être retenue que si un problème de compatibilité ou de performance avéré l'impose.

## Points de vigilance

Plusieurs précautions méritent d'être gravées dans les réflexes.

La plus importante concerne l'irréversibilité en cas d'oubli : il n'existe aucun mécanisme officiel de récupération d'un mot de passe perdu. Microsoft ne peut pas le retrouver, et avec le chiffrement fort des versions récentes, une base dont on a perdu le mot de passe est, à toutes fins pratiques, définitivement inaccessible. Conservez donc le mot de passe dans un gestionnaire sûr, et réalisez systématiquement une sauvegarde de la base avant de la chiffrer.

En environnement multi-utilisateur, gardez à l'esprit qu'un changement du mot de passe du back-end impose de reconnecter toutes les tables liées de tous les front-ends déployés. Cette opération doit être planifiée et, idéalement, automatisée, sous peine de bloquer l'ensemble des postes.

Enfin, n'oubliez pas que cette protection est complémentaire des autres. Le mot de passe chiffre le fichier mais ne protège ni le code (rôle du format `.accde`, voir [section 20.3](03-compilation-accde.md)) ni les droits fonctionnels des utilisateurs à l'intérieur de l'application (rôle d'une gestion applicative, voir [section 20.6](06-gestion-droits-utilisateurs.md)). Le chiffrement par mot de passe et la compilation en `.accde` sont d'ailleurs parfaitement cumulables sur un même fichier.

## Ce qu'il faut retenir

Le mot de passe de base de données est un mécanisme simple mais puissant lorsqu'il est bien employé : il chiffre tout le fichier et, sur une version récente d'Access avec un mot de passe fort, offre une protection réelle des données au repos. Sa pose exige une ouverture en exclusif et s'automatise en DAO via `NewPassword`. Ses principaux pièges sont l'absence totale de récupération en cas d'oubli, l'exposition du mot de passe dans les tables liées d'un front-end, et la nécessité de choisir un mot de passe robuste. Manié avec ces précautions, il constitue le socle sur lequel viendront s'ajouter les protections du code et la sécurité applicative décrites dans la suite du chapitre.

---


⏭️ [20.3. Compilation en ACCDE — verrouillage du code VBA](/20-securite-protection/03-compilation-accde.md)
