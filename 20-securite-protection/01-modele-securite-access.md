🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 20.1. Modèle de sécurité Access — historique et situation actuelle

Avant d'examiner les protections concrètes présentées dans les sections suivantes, il est essentiel de comprendre ce qu'est — et surtout ce que n'est plus — le modèle de sécurité d'Access. Ce modèle a connu une rupture majeure au milieu des années 2000, et beaucoup d'idées reçues qui circulent encore aujourd'hui datent de l'ancien système, disparu depuis longtemps. Comprendre cette histoire permet d'éviter deux erreurs symétriques : chercher des fonctionnalités qui n'existent plus, et mal interpréter ce que les protections actuelles garantissent réellement.

## L'héritage : la sécurité au niveau utilisateur des fichiers MDB

Jusqu'à la version 2003 incluse, Access reposait sur le moteur Jet et sur le format de fichier `.mdb`. Ce format proposait un dispositif appelé sécurité au niveau utilisateur (en anglais *User-Level Security*, souvent abrégé ULS). Il s'agissait d'un véritable système de comptes : on définissait des utilisateurs et des groupes, puis on attribuait des permissions précises sur chaque objet de la base (ouvrir, lire, modifier, supprimer une table, exécuter une requête, ouvrir un formulaire, et ainsi de suite).

Ces informations d'identité n'étaient pas stockées dans la base elle-même mais dans un fichier séparé appelé fichier de groupe de travail, portant l'extension `.mdw` (le fichier par défaut se nommait `System.mdw`). Pour ouvrir une base sécurisée, Access devait être associé au bon fichier de groupe, qui contenait la définition des utilisateurs et leurs mots de passe. Deux groupes existaient toujours par défaut, `Admins` et `Users`, ainsi qu'un compte administrateur nommé `Admin`.

Sur le papier, ce système était séduisant car il offrait une granularité fine, proche de celle d'un vrai serveur de bases de données. En pratique, il souffrait de défauts rédhibitoires. Sa configuration était notoirement complexe et déroutante : une erreur de manipulation, comme oublier de retirer les permissions du groupe `Users` ou du compte `Admin`, suffisait à rendre l'ensemble inopérant. La faille la plus connue tenait justement à ce compte `Admin`, doté par défaut d'un mot de passe vide : tant qu'on ne l'avait pas explicitement neutralisé, n'importe qui pouvait s'en servir. Surtout, le chiffrement sous-jacent était faible, au point que des outils de récupération gratuits permettaient de contourner la protection en quelques instants. La sécurité au niveau utilisateur donnait donc un sentiment de protection bien supérieur à la réalité.

## La rupture de 2007 : le passage à ACCDB

Avec Access 2007, Microsoft a introduit un nouveau moteur, ACE (*Access Connectivity Engine*), et un nouveau format de fichier, `.accdb`, appelé à remplacer le `.mdb`. Ce changement s'est accompagné d'une décision lourde de conséquences : la sécurité au niveau utilisateur a été purement et simplement supprimée pour le nouveau format. Les notions de comptes, de groupes, de permissions par objet et de fichier de groupe de travail ne s'appliquent plus du tout aux bases `.accdb`.

Cette suppression n'était pas un oubli mais un choix assumé. Microsoft a considéré que l'ancien système était à la fois trop complexe pour être correctement utilisé par la plupart des développeurs et trop faible pour offrir une protection digne de ce nom. Plutôt que de maintenir une fonctionnalité qui induisait les utilisateurs en erreur, l'éditeur a préféré la retirer et réorienter la sécurité d'Access autour d'autres mécanismes. À noter : la sécurité au niveau utilisateur reste techniquement disponible si l'on continue d'employer l'ancien format `.mdb`, mais ce format est aujourd'hui considéré comme obsolète et ne doit plus servir à de nouveaux développements.

## La situation actuelle : sur quoi repose la sécurité d'une base ACCDB

Dans le modèle actuel, la sécurité d'une application Access ne repose plus sur un dispositif unique et intégré, mais sur la combinaison de quatre éléments distincts, qui constituent précisément le plan de ce chapitre.

Le premier est le chiffrement du fichier par mot de passe. Lorsqu'on définit un mot de passe de base de données sur un fichier `.accdb`, Access chiffre l'intégralité du contenu. Il s'agit cette fois d'un véritable chiffrement, et non du simple encodage trivial des anciens `.mdb`. À partir d'Access 2010, ce chiffrement s'appuie sur l'API cryptographique de Windows et sur un algorithme moderne, ce qui le rend nettement plus solide que l'implémentation initiale d'Access 2007 et sans commune mesure avec le format `.mdb`. Il faut toutefois bien comprendre la nature de cette protection : c'est un mot de passe unique, partagé par tous les utilisateurs. Il commande l'ouverture du fichier de manière binaire — soit on connaît le mot de passe et on accède à tout, soit on l'ignore et on n'accède à rien. Ce n'est donc pas un système de comptes individuels. Ce mécanisme est détaillé en [section 20.2](02-protection-mot-de-passe.md).

Le deuxième élément est la protection du code. Le format `.accde`, obtenu par compilation de la base, supprime purement et simplement le code source VBA tout en conservant les versions compilées et exécutables des modules. C'est aujourd'hui la méthode robuste pour empêcher la lecture et la modification du code de l'application. Ce point est traité en [section 20.3](03-compilation-accde.md).

Le troisième élément relève de la sécurité au niveau d'Office tout entier, et non d'Access seul. Le centre de gestion de la confidentialité régit la confiance accordée au code à l'ouverture : emplacements approuvés, niveaux de sécurité des macros, barre des messages avertissant l'utilisateur. La signature numérique du projet, abordée en [section 20.4](04-signature-numerique.md), s'inscrit dans ce cadre, tout comme les réglages présentés en [section 20.5](05-parametres-securite-macro.md).

Le quatrième et dernier élément est le plus important à intégrer mentalement : la sécurité applicative est désormais à la charge du développeur. Puisqu'Access ne fournit plus de gestion des droits par utilisateur, c'est à vous de la construire si votre application en a besoin, typiquement au moyen d'une table de rôles couplée à du code qui adapte l'interface selon le profil connecté. Cette approche fait l'objet de la [section 20.6](06-gestion-droits-utilisateurs.md), complétée par l'audit des actions en [section 20.7](07-audit-acces-modifications.md). Il s'agit d'une sécurité fonctionnelle, qui contrôle ce qu'un utilisateur a le droit de faire à l'intérieur de l'application, mais qui ne protège pas le fichier sur le disque : elle suppose donc d'être combinée avec les protections de fichier précédentes.

## Le cas particulier du mot de passe de projet VBA

Indépendamment du format de fichier, l'éditeur VBA permet de protéger l'affichage d'un projet par un mot de passe, via les propriétés du projet et leur onglet de protection. Il est important de ne pas confondre ce dispositif avec la compilation en `.accde`. Le mot de passe de projet VBA empêche seulement l'ouverture du code dans l'éditeur, mais le code source reste présent dans le fichier ; or des outils répandus permettent de le retirer très facilement. Cette protection ne doit donc être vue que comme un garde-fou minimal contre la curiosité, jamais comme une véritable sécurité. Pour protéger sérieusement votre code, c'est le format `.accde` qu'il faut employer.

## Que reste-t-il de DAO côté sécurité ?

Pour des raisons de compatibilité, la bibliothèque DAO expose toujours des collections `Users` et `Groups` au niveau de l'objet `Workspace`, ainsi qu'une propriété `Permissions` sur les objets. Il ne faut pas se laisser abuser par leur présence : ces éléments sont des vestiges de l'ancien modèle et n'ont plus d'effet de sécurité réel sur une base `.accdb`. Tenter de les manipuler ne reconstitue en aucun cas la sécurité au niveau utilisateur d'autrefois.

```vba
' Sur une base .accdb, ces collections existent encore par héritage
' mais ne procurent AUCUNE sécurité effective.
Dim ws As DAO.Workspace
Set ws = DBEngine.Workspaces(0)

Debug.Print "Utilisateur courant : " & ws.UserName  ' renverra "Admin"

' Ne vous appuyez jamais là-dessus pour protéger une base .accdb :
' la véritable sécurité passe par le chiffrement du fichier,
' le format .accde et une gestion applicative des droits.
Set ws = Nothing
```

## Synthèse : MDB hier, ACCDB aujourd'hui

Le tableau suivant résume le glissement d'un modèle à l'autre.

| Aspect | Fichiers `.mdb` (ancien modèle) | Fichiers `.accdb` (modèle actuel) |
| --- | --- | --- |
| Comptes et groupes d'utilisateurs | Oui, via un fichier de groupe `.mdw` | Supprimés |
| Permissions par objet | Oui, au niveau du moteur | Non disponibles |
| Mot de passe de base | Oui, mais encodage très faible | Oui, chiffrement réel (renforcé depuis Access 2010) |
| Nature du mot de passe | Combiné à des comptes individuels | Unique et partagé par tous |
| Protection du code | Mot de passe de projet VBA (faible) | Format `.accde` (suppression du source) |
| Gestion des droits métier | Assurée par la sécurité native | À développer soi-même (table de rôles) |

## Ce qu'il faut retenir

La sécurité d'Access a changé de nature. L'ancien système de comptes et de permissions intégrés n'existe plus dans les fichiers `.accdb` : toute documentation ou tout tutoriel évoquant les fichiers de groupe de travail, les utilisateurs Jet ou les permissions par objet décrit un monde révolu. Aujourd'hui, protéger une application Access consiste à empiler des mesures complémentaires — chiffrer le fichier, verrouiller le code, configurer la confiance d'Office et bâtir une gestion applicative des droits — tout en gardant à l'esprit qu'aucune de ces mesures ne transforme Access en une forteresse. Pour des données réellement sensibles, le bon réflexe demeure de déléguer le stockage à un moteur conçu pour la sécurité, comme exposé au [chapitre 23](../23-migration-interoperabilite/README.md). Les sections qui suivent détaillent une à une ces différentes couches de protection.

---


⏭️ [20.2. Protection par mot de passe de la base de données](/20-securite-protection/02-protection-mot-de-passe.md)
