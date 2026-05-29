🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 23.4. Architecture hybride Access + SQL Server (tables liées ODBC)

Les sections précédentes ont conduit les données vers SQL Server (23.2) et adapté le code en conséquence (23.3). Il reste à décrire l'architecture qui en résulte et dans laquelle l'application va vivre durablement : un **front-end Access** connecté à un **back-end SQL Server** par des tables liées ODBC. C'est de très loin la cible de migration la plus répandue, parce qu'elle conjugue deux avantages rarement réunis — préserver l'intégralité de l'investissement réalisé dans l'interface Access, tout en confiant les données à un véritable moteur serveur. Cette section explique comment la mettre en place, comment la rendre fiable, et quelles en sont les forces comme les limites.

## Le principe : séparer le front-end du back-end

L'idée fondatrice est la même que celle de la base scindée vue au chapitre 15, mais le back-end n'est plus un fichier Access : c'est SQL Server. Le **front-end** est un fichier Access (idéalement compilé en `.accde`) qui contient toute l'application — formulaires, états, requêtes, modules VBA, ruban — mais aucune donnée. Le **back-end** est l'instance SQL Server qui héberge les tables, les index, les vues et éventuellement les procédures stockées. Entre les deux, des tables liées par ODBC font le pont : dans le front-end, ces liens apparaissent comme des tables ordinaires, alors que les enregistrements résident en réalité sur le serveur.

```
   Poste utilisateur 1        Poste utilisateur 2        Poste utilisateur N
  ┌──────────────────┐       ┌──────────────────┐       ┌──────────────────┐
  │ Front-end Access │       │ Front-end Access │       │ Front-end Access │
  │ (.accde)         │       │ (.accde)         │  ...  │ (.accde)         │
  │ formulaires,     │       │ formulaires,     │       │ formulaires,     │
  │ états, VBA       │       │ états, VBA       │       │ états, VBA       │
  └────────┬─────────┘       └────────┬─────────┘       └────────┬─────────┘
           │                          │                          │
           └─────────── ODBC ─────────┼─────────── ODBC ─────────┘
                                      │
                             ┌────────▼─────────┐
                             │   SQL Server     │
                             │   (back-end)     │
                             │ tables, données, │
                             │ index, vues,     │
                             │ procédures       │
                             └──────────────────┘
```

Cette répartition explique le succès du modèle. L'utilisateur retrouve exactement son interface habituelle ; aucune formation n'est nécessaire. Le développeur conserve son code VBA, ses formulaires et ses états. Et la couche de données gagne tout ce qui manquait à Access : robustesse, sécurité, sauvegarde à chaud, gestion fine de la concurrence et capacité de montée en charge. La migration devient ainsi progressive et peu risquée, puisqu'on ne remplace qu'une couche à la fois.

## La liaison par ODBC

### Le pilote ODBC

Pour qu'un poste se connecte au serveur, il doit disposer du **pilote ODBC pour SQL Server**. Microsoft fournit des pilotes dédiés (par exemple *ODBC Driver 17 for SQL Server*, ou une version plus récente) qu'il convient d'installer sur chaque poste client et de garder homogènes dans tout le parc. Le nom exact du pilote figure dans la chaîne de connexion, il faut donc s'assurer que le pilote référencé est bien celui installé sur les postes.

### DSN ou connexion sans DSN

Deux manières de définir la connexion coexistent, et le choix a un fort impact sur le déploiement.

| Approche | Où est définie la connexion | Conséquence pour le déploiement |
|---|---|---|
| DSN (User, System ou File) | Dans une source de données configurée via l'administrateur ODBC de Windows | À recréer sur **chaque poste** : lourd à maintenir |
| Sans DSN (DSN-less) | Directement dans la propriété `Connect` de la table liée | Rien à configurer par poste : recommandé |

La connexion **sans DSN** est presque toujours préférable : la chaîne de connexion étant stockée dans la définition de la table liée elle-même, il n'y a rien à paramétrer sur les postes, ce qui simplifie considérablement la distribution et la maintenance. Les formats de chaîne de connexion sont détaillés à l'annexe D.

### Lier les tables par code

Plutôt que de passer par l'assistant graphique (Données externes → Base de données ODBC), une application robuste crée et rafraîchit ses liens **par code**, ce qui permet de les recréer automatiquement lors du déploiement. En DAO, on construit un `TableDef` dont on renseigne la table source côté serveur et la chaîne de connexion (voir aussi section 12.7).

```vba
Public Sub LierTableSQL(strTableServeur As String, strAliasLocal As String)
    Dim db As DAO.Database
    Dim tdf As DAO.TableDef
    Set db = CurrentDb

    ' Supprimer un éventuel lien existant portant le même nom
    On Error Resume Next
    db.TableDefs.Delete strAliasLocal
    On Error GoTo 0

    Set tdf = db.CreateTableDef(strAliasLocal)
    tdf.SourceTableName = strTableServeur
    ' Chaîne de connexion sans DSN (authentification Windows)
    tdf.Connect = "ODBC;Driver={ODBC Driver 17 for SQL Server};" & _
                  "Server=SRV-SQL\INSTANCE;Database=GestionVentes;" & _
                  "Trusted_Connection=Yes;"
    db.TableDefs.Append tdf
    db.TableDefs.Refresh
End Sub
```

## Reliaison et gestion des connexions

Les tables liées doivent être reconnectées dans plusieurs situations : déploiement du front-end sur un nouveau poste, passage de l'environnement de développement à la production, ou changement de nom du serveur. Une application bien conçue ne demande jamais à l'utilisateur de reconfigurer manuellement ses liens ; elle les rétablit elle-même, typiquement au démarrage.

Une routine de reliaison parcourt les tables liées ODBC et met à jour leur connexion.

```vba
Public Sub ReconnecterTables(strNouvelleConnexion As String)
    Dim db As DAO.Database
    Dim tdf As DAO.TableDef
    Set db = CurrentDb

    For Each tdf In db.TableDefs
        ' Ne traiter que les tables liées ODBC
        If (tdf.Attributes And dbAttachedODBC) <> 0 Then
            tdf.Connect = strNouvelleConnexion
            tdf.RefreshLink
        End If
    Next tdf
End Sub
```

Pour que cette reliaison soit pilotable, les informations de connexion (nom du serveur, base, mode d'authentification) gagnent à être **centralisées** — dans une table de configuration, un fichier de paramètres, le registre ou des `TempVars` — plutôt que codées en dur. L'application lit ces paramètres au lancement et reconnecte ses tables si nécessaire. Combinée aux connexions sans DSN, cette approche rend le déploiement quasi automatique et constitue la base des stratégies de mise à jour et de reliaison automatique vues au chapitre 21 (sections 21.3 et 21.4).

## Les conditions d'une liaison qui fonctionne bien

Une table liée mal préparée côté serveur peut s'avérer lente, en lecture seule, ou afficher des erreurs déconcertantes. Quelques règles évitent l'essentiel des problèmes.

### Clé primaire ou index unique

C'est la condition la plus importante. Pour qu'Access considère une table ou une vue liée comme **modifiable**, il lui faut un moyen d'identifier chaque ligne de façon unique. Une table dotée d'une **clé primaire** ne pose pas de problème. En revanche, une table sans clé primaire, ou surtout une **vue**, sera en lecture seule tant qu'Access ne connaît pas un index unique : au moment de lier une vue, Access demande quel champ (ou groupe de champs) l'identifie de manière unique, et passer cette étape donne une vue non modifiable. Beaucoup de problèmes de « données impossibles à modifier » après migration viennent simplement d'une clé manquante.

### Une colonne rowversion pour la détection des modifications

Ajouter une colonne de type `rowversion` (anciennement `timestamp`) à chaque table SQL Server est une excellente pratique. Access s'en sert pour détecter efficacement si une ligne a changé depuis sa lecture, ce qui fiabilise la concurrence optimiste et évite les faux conflits d'écriture, notamment sur les tables comportant des nombres à virgule flottante ou des champs `bit`.

### Champs bit : NOT NULL et valeur par défaut

Les colonnes booléennes (`bit`) doivent de préférence être déclarées **NOT NULL avec une valeur par défaut** (généralement 0). Un `bit` à NULL perturbe Access, qui peut alors mal afficher l'enregistrement. Définir `bit NOT NULL DEFAULT 0` évite cette catégorie de problèmes.

### Comprendre « #Deleted » et les conflits d'écriture

L'affichage de `#Deleted` (ou `#Supprimé`) à la place des données, ou des messages indiquant qu'« un autre utilisateur a modifié l'enregistrement » alors que personne ne l'a touché, sont des symptômes classiques de table liée mal configurée. Les causes sont presque toujours dans la liste suivante.

| Cause | Correctif |
|---|---|
| Absence de clé primaire ou d'index unique | Ajouter une clé primaire / déclarer un index unique au moment de la liaison |
| Colonne `bit` autorisant NULL | Passer la colonne en `NOT NULL DEFAULT 0` |
| Absence de colonne `rowversion` | Ajouter une colonne `rowversion` à la table |
| Arrondis sur des colonnes à virgule flottante | Préférer `decimal` ; la colonne `rowversion` neutralise aussi le problème |
| Déclencheur (trigger) modifiant la ligne insérée | Revoir le déclencheur ou utiliser `rowversion` pour qu'Access détecte correctement l'état |

## Performance dans l'architecture hybride

Relier les tables ne suffit pas à garantir de bonnes performances ; encore faut-il solliciter le serveur intelligemment, dans l'esprit déjà exposé à la section 23.3.

Le point le plus sensible reste les **formulaires liés à des tables entières** : un formulaire dont la source est une table complète charge l'intégralité des enregistrements depuis le serveur à son ouverture. Il faut restreindre la source par un filtre ou une requête paramétrée et n'afficher que ce dont l'utilisateur a besoin (section 18.5). Le même soin s'applique aux **listes déroulantes et zones de liste** branchées sur de grosses tables, qu'il convient de limiter par une clause `WHERE`.

Pour les **traitements lourds et les états**, les requêtes Pass-Through (section 11.9) et les **vues** côté serveur permettent de pré-filtrer et pré-joindre les données avant de les rapatrier, en exploitant pleinement la puissance du moteur.

Enfin, une optimisation discrète mais très efficace consiste à **maintenir une connexion ouverte** pendant toute la session. Sans cela, Access établit et ferme des connexions ODBC à répétition, ce qui est coûteux. Garder un Recordset léger ouvert en arrière-plan conserve la connexion active et accélère sensiblement les opérations suivantes (section 18.8).

```vba
' Au démarrage : ouvrir un Recordset léger et le conserver ouvert
' pour maintenir la connexion ODBC active toute la session
Private mrsKeepAlive As DAO.Recordset

Public Sub OuvrirConnexionPersistante()
    Set mrsKeepAlive = CurrentDb.OpenRecordset( _
        "SELECT TOP 1 1 AS X FROM Clients", dbOpenSnapshot)
End Sub
```

## Distribution et maintenance du front-end

Dans une architecture hybride, **chaque utilisateur dispose de sa propre copie du front-end**, installée localement sur son poste, et toutes ces copies pointent vers le même serveur. Partager un unique fichier front-end entre plusieurs utilisateurs réintroduirait, pour les objets de l'application, les problèmes de fichier partagé que la migration cherchait justement à éliminer.

Le front-end est de préférence distribué sous forme de `.accde`, ce qui protège le code et empêche les modifications accidentelles (section 20.3). Sa mise à jour et la reliaison automatique de ses tables relèvent des stratégies de déploiement du chapitre 21. Côté serveur, l'administration (sauvegardes, sécurité, supervision) devient une responsabilité distincte, prise en charge avec les outils de SQL Server.

## Forces et limites de l'architecture hybride

Les forces de ce modèle sont considérables : il préserve l'investissement applicatif, sécurise et fiabilise la couche de données, autorise une migration graduelle et n'impose aucun changement aux utilisateurs. Pour une application Access devenue critique mais dont l'interface donne satisfaction, c'est presque toujours la meilleure trajectoire.

Ses limites tiennent à ce qu'il reste, par construction, une application Access. Le front-end doit être installé et mis à jour sur chaque poste, et l'on demeure dans un environnement bureautique de bureau, sans accès web ni mobile natif. La couche ODBC ajoute un léger surcoût, et la latence réseau jusqu'au serveur continue de peser sur les opérations trop « bavardes ». Surtout, l'architecture hybride ne répond pas aux besoins d'accès en mobilité ou par navigateur ; ces scénarios appellent d'autres solutions, examinées dans les sections suivantes.

## En résumé

L'architecture hybride associe un front-end Access — formulaires, états et VBA inchangés — à un back-end SQL Server, via des tables liées ODBC de préférence sans DSN et reconnectées par code au démarrage. Sa fiabilité repose sur quelques conditions précises côté serveur : clé primaire ou index unique, colonne `rowversion`, champs `bit` non nuls par défaut. Ses performances dépendent d'un usage mesuré du serveur : formulaires filtrés, vues, requêtes Pass-Through et connexion maintenue ouverte. C'est le débouché naturel d'une migration vers SQL Server lorsque l'on souhaite conserver Access. Lorsque ce cadre ne suffit plus — besoin de web, de mobilité, ou d'intégration à l'écosystème Microsoft 365 — il faut envisager une migration vers d'autres plateformes, à commencer par SharePoint et Dataverse, objet de la section suivante.

⏭️ [23.5. Migration vers SharePoint Lists et Dataverse](/23-migration-interoperabilite/05-migration-sharepoint-dataverse.md)
