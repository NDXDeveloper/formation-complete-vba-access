🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 15.1. Architecture multi-utilisateurs — front-end / back-end

Avant toute considération sur le verrouillage ou les conflits, il faut poser la fondation sur laquelle repose tout le multi-utilisateur sous Access : la **séparation de l'application en deux fichiers distincts**, un front-end et un back-end. Cette architecture n'est pas une simple recommandation parmi d'autres ; c'est la condition d'une application partagée fiable. Cette section en explique le principe, les raisons, le déploiement et les bonnes pratiques.

## Le principe : séparer l'interface des données

L'idée est de scinder l'application Access monolithique habituelle en deux fichiers aux rôles bien distincts.

Le **back-end** (BE) contient uniquement les **données** : les tables et les relations entre elles. C'est l'unique source de vérité, et il est stocké à un emplacement central, partagé sur le réseau. Il n'y a qu'**un seul** back-end.

Le **front-end** (FE) contient **tout le reste** : les formulaires, les états, les requêtes, les macros et le code VBA — ainsi que des **tables liées** qui pointent vers les tables réelles situées dans le back-end. Point crucial : **chaque utilisateur dispose de sa propre copie du front-end**, installée localement sur son poste.

On obtient ainsi une étoile : un back-end central de données, vers lequel pointent plusieurs front-ends locaux, un par utilisateur.

## Pourquoi séparer ?

Cette séparation répond à plusieurs exigences simultanées.

**La concurrence.** Plusieurs utilisateurs travaillent en même temps, chacun dans son propre front-end. Leurs objets locaux (formulaires ouverts, requêtes temporaires, état de l'interface) ne se télescopent pas, tandis que tous lisent et écrivent dans le même back-end partagé. C'est ce qui rend l'accès multi-utilisateur viable.

**La résilience à la corruption.** Le risque de corruption d'un fichier Access est concentré là où se produit l'essentiel de l'activité — l'ouverture de formulaires, la création de recordsets, les objets temporaires. En isolant cette activité dans le front-end, on protège les données : si un front-end se corrompt, seul l'utilisateur concerné est affecté, et l'on remplace simplement sa copie. Le back-end, lui, ne contient que des données et subit beaucoup moins de sollicitations structurelles.

**La performance.** Les formulaires, états et requêtes s'exécutent localement, à partir du front-end installé sur le poste. Seules les données transitent par le réseau, ce qui réduit considérablement le trafic par rapport à un fichier unique où toute l'interface circulerait elle aussi. L'optimisation réseau des bases partagées est approfondie à la [section 18.8](/18-optimisation-performance/08-optimisation-reseau.md).

**La maintenance et le déploiement.** C'est un avantage décisif : on peut corriger un bug ou ajouter une fonctionnalité dans le front-end et redéployer la nouvelle version aux utilisateurs **sans jamais toucher aux données**. Le back-end continue de fonctionner, les données restent en place. Mettre à jour une application monolithique imposerait au contraire d'extraire puis de réimporter les données à chaque livraison. Les stratégies de mise à jour du front-end sont traitées à la [section 21.3](/21-deploiement-distribution/03-mise-a-jour-front-end.md).

**La sauvegarde.** Sauvegarder le back-end suffit à protéger ce qui compte vraiment — les données. Les front-ends, eux, sont reproductibles à partir d'une copie maîtresse.

## Le piège du fichier unique partagé

À l'opposé de cette architecture se trouve la solution naïve : un **fichier unique** contenant à la fois les données et l'interface, ouvert directement par tous les utilisateurs depuis le réseau. Cette approche cumule les inconvénients.

Tous les utilisateurs sollicitant le même fichier, et toute l'activité de l'interface s'y déroulant par le réseau, le **risque de corruption est nettement plus élevé**. Il devient impossible de mettre à jour l'application sans en exclure tout le monde. Le trafic réseau explose, et la contention sur les verrous augmente. Cette configuration n'est acceptable que pour un usage strictement mono-utilisateur ou un partage occasionnel et léger ; pour toute véritable application partagée, elle est à proscrire.

## Les tables liées : le pont entre front-end et back-end

Le lien entre les deux fichiers est assuré par les **tables liées** du front-end. Une table liée n'est pas une copie des données : c'est un objet de référence qui pointe vers une table réelle du back-end. Elle apparaît dans le front-end et se manipule — dans les formulaires, les requêtes, le code — presque comme une table locale, mais les données qu'elle expose résident physiquement dans le back-end.

Techniquement, chaque table liée conserve dans sa définition le **chemin du back-end** et le **nom de la table source**. On peut inspecter ces informations par code (voir l'illustration plus bas), et surtout créer ou réparer ces liaisons par programmation — opération essentielle pour le déploiement, détaillée à la [section 15.9](09-tables-liees-par-code.md).

## Le modèle de déploiement

Deux règles gouvernent un déploiement correct.

Le **back-end est placé sur un partage réseau** accessible à tous les postes. Il est fortement recommandé d'y accéder par un chemin UNC (de la forme `\\Serveur\Partage\...`) plutôt que par une lettre de lecteur mappée, dont la disponibilité varie d'un poste à l'autre. Cet aspect est traité à la [section 21.5](/21-deploiement-distribution/05-deploiement-reseau-chemins-unc.md).

**Chaque utilisateur exécute sa propre copie locale du front-end.** C'est la règle la plus importante — et l'erreur la plus fréquente consiste à la transgresser. Faire ouvrir un même front-end situé sur le réseau par plusieurs utilisateurs réintroduit exactement les problèmes du fichier unique : corruption, contention, trafic. Le front-end doit être copié sur le disque local de chaque poste et exécuté de là.

## Scinder une base existante

Access fournit un assistant intégré, le **« Fractionneur de base de données »** (Outils de base de données → Access Database / Déplacer les données), qui automatise la séparation : il déplace les tables vers un nouveau fichier back-end et remplace les tables locales du fichier courant par des tables liées pointant vers ce back-end. C'est le moyen le plus simple de transformer une application monolithique existante en architecture front-end / back-end.

La séparation peut aussi être réalisée manuellement, ou pilotée par code lorsqu'on industrialise le déploiement — la création des liaisons par VBA faisant l'objet de la [section 15.9](09-tables-liees-par-code.md).

## Que placer où ?

La répartition des objets découle directement des rôles :

- **dans le back-end** : les tables et les relations entre elles. C'est aussi là que vit l'intégrité référentielle (voir [chapitre 14](/14-transactions/06-integrite-referentielle.md)), puisque les règles accompagnent les données qu'elles protègent ;
- **dans le front-end** : les formulaires, les états, les requêtes, les macros, les modules de code, et les tables liées.

Une nuance utile : certaines tables peuvent légitimement rester **locales au front-end**. C'est typiquement le cas des tables de travail temporaires propres à un traitement individuel, ou de paramètres spécifiques à un poste. Les garder dans le front-end évite toute contention sur le back-end et améliore les performances — c'est un bénéfice indirect de l'architecture, et non une entorse à la règle.

## Illustration : inspecter les tables liées

Le court extrait suivant parcourt les tables d'un front-end et affiche, pour chaque table liée, sa source. Une table liée se reconnaît à sa propriété `Connect` non vide ; `SourceTableName` indique le nom de la table dans le back-end. (La création et la réparation de ces liaisons sont traitées en détail à la [section 15.9](09-tables-liees-par-code.md).)

```vba
Public Sub ListerTablesLiees()
    Dim db As DAO.Database
    Dim tdf As DAO.TableDef

    Set db = CurrentDb
    For Each tdf In db.TableDefs
        ' Une table liée possède une chaîne Connect non vide
        If Len(tdf.Connect) > 0 Then
            Debug.Print tdf.Name & "  ->  source : " & tdf.SourceTableName
            Debug.Print "      Connect : " & tdf.Connect
        End If
    Next tdf

    Set db = Nothing
End Sub
```

Pour une table liée à un back-end Access, la chaîne `Connect` se présente typiquement sous la forme `;DATABASE=\\Serveur\Partage\MaBase_be.accdb` ; pour une source externe via ODBC, elle commence par `ODBC;` (cf. [chapitre 23](/23-migration-interoperabilite/README.md)).

## Points de vigilance

- **Un seul back-end, autant de front-ends que d'utilisateurs.** C'est le schéma de référence du multi-utilisateur Access.
- **Jamais de front-end partagé sur le réseau.** Chaque poste exécute sa copie locale ; l'oublier annule tous les bénéfices de la séparation.
- **Le back-end ne contient que des données.** Aucun formulaire, état ou module n'y a sa place.
- **Privilégier les chemins UNC** pour le back-end, plus robustes que les lettres de lecteur mappées (section 21.5).
- **Penser la reliaison dès le déploiement.** Le chemin du back-end peut changer d'un environnement à l'autre ; prévoir une reliaison automatique des tables (sections 15.9 et 21.4) évite bien des incidents.
- **Surveiller la taille du back-end.** Toutes les données étant concentrées dans un fichier, la limite de 2 Go s'y applique (section 18.9).

## En résumé

L'architecture multi-utilisateur d'Access repose sur la séparation en un **back-end** central, qui ne contient que les tables et les relations, et plusieurs **front-ends** locaux — un par utilisateur — qui contiennent l'interface (formulaires, états, requêtes, code) et des **tables liées** pointant vers le back-end. Cette séparation est la condition de la concurrence, de la résilience à la corruption, de la performance et d'une maintenance sans interruption de service. Le fichier unique partagé, à l'inverse, est à proscrire pour toute application réellement partagée. Côté déploiement, deux règles priment : un back-end sur partage réseau (chemin UNC), et une copie locale du front-end par poste. La mécanique de liaison des tables qui rend tout cela possible est l'objet de la section 15.9.

La section suivante précise la manière dont une base est ouverte par plusieurs utilisateurs : les [modes de partage de la base de données](02-modes-partage.md).

⏭️ [15.2. Modes de partage de la base de données](/15-multi-utilisateurs/02-modes-partage.md)
