🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 15.9. Connexion à une base distante via table liée par code

L'architecture front-end / back-end de la [section 15.1](01-architecture-front-end-back-end.md) repose sur les **tables liées** : des objets du front-end qui pointent vers les tables réelles du back-end. Or le chemin du back-end change souvent — d'un environnement de développement à la production, d'un partage réseau à un autre, d'un poste de test à un serveur. Savoir créer et surtout **relier** ces tables par code est donc indispensable au déploiement et à la maintenance d'une application partagée. Cette section détaille ces opérations.

## Rappel : qu'est-ce qu'une table liée ?

Une table liée n'est pas une copie des données : c'est une référence vers une table d'une autre base. En DAO, elle est représentée par un objet `TableDef` dont deux propriétés portent le lien :

- **`Connect`** : la chaîne de connexion vers la base source. Une chaîne non vide est précisément ce qui distingue une table liée d'une table locale (cf. l'inspection vue en 15.1).
- **`SourceTableName`** : le nom de la table telle qu'elle existe dans la base source.

Manipuler ces deux propriétés, plus la méthode `RefreshLink`, suffit à gérer entièrement les liaisons par programmation.

## Créer une table liée par code

La création se fait avec `CreateTableDef`, en renseignant `SourceTableName` et `Connect`, puis en ajoutant l'objet à la collection `TableDefs`.

```vba
Public Sub LierTable(nomLien As String, nomSource As String, cheminBE As String)
    Dim db As DAO.Database
    Dim tdf As DAO.TableDef

    Set db = CurrentDb

    ' Supprimer un éventuel lien existant de même nom
    On Error Resume Next
    db.TableDefs.Delete nomLien
    On Error GoTo 0

    Set tdf = db.CreateTableDef(nomLien)        ' nom du lien dans le front-end
    tdf.SourceTableName = nomSource             ' nom de la table dans le back-end
    tdf.Connect = ";DATABASE=" & cheminBE       ' back-end Access
    db.TableDefs.Append tdf                     ' crée le lien

    Set tdf = Nothing: Set db = Nothing
End Sub
```

Pour un back-end Access, la chaîne `Connect` prend la forme `;DATABASE=<chemin>` — **avec un point-virgule en tête** et sans préfixe de fournisseur (c'est une caractéristique souvent oubliée). Si le back-end est protégé par mot de passe, on ajoute `;PWD=<motdepasse>`.

Le nom du lien (`nomLien`) est celui sous lequel la table apparaîtra dans le front-end ; il peut différer de `SourceTableName`, mais on les garde généralement identiques pour la lisibilité.

## Lier une source ODBC

Pour une source externe — SQL Server, par exemple —, seule la chaîne `Connect` change : elle commence par `ODBC;` et décrit le pilote et le serveur.

```vba
Set tdf = db.CreateTableDef("Clients")
tdf.SourceTableName = "dbo.Clients"             ' schéma + table côté serveur
tdf.Connect = "ODBC;DRIVER={ODBC Driver 17 for SQL Server};" & _
              "SERVER=MonServeur;DATABASE=MaBase;Trusted_Connection=Yes;"
' tdf.Attributes = tdf.Attributes Or dbAttachSavePWD   ' enregistre les identifiants
db.TableDefs.Append tdf
```

Pour une table de serveur, on précise généralement le schéma dans `SourceTableName` (`dbo.Clients`). L'attribut **`dbAttachSavePWD`** permet d'enregistrer les identifiants dans le lien, évitant à l'utilisateur de les ressaisir — au prix d'un compromis de sécurité, car le mot de passe est alors stocké dans le front-end (à mettre en regard du [chapitre 20](/20-securite-protection/README.md)). La liaison à des serveurs externes est approfondie au [chapitre 23](/23-migration-interoperabilite/README.md).

## L'alternative DoCmd.TransferDatabase

On peut aussi créer un lien sans manipuler directement les `TableDef`, au moyen de `DoCmd.TransferDatabase` avec l'argument `acLink` :

```vba
DoCmd.TransferDatabase acLink, "Microsoft Access", _
    "\\Serveur\Partage\MaBase_be.accdb", acTable, "Clients", "Clients"
```

Cette méthode (vue côté import/export à la [section 5.5](/05-objet-docmd/05-import-export-donnees.md)) convient pour une création initiale. Pour la **reliaison**, l'approche par `Connect` + `RefreshLink` ci-dessous offre plus de contrôle et reste la référence.

## Reliaison : changer le chemin du back-end

C'est l'opération la plus importante du déploiement. Pour repointer des liens existants vers un nouveau back-end, on met à jour la propriété `Connect` de chaque table liée, puis on appelle `RefreshLink`, qui rétablit la connexion à partir de la chaîne courante.

```vba
Public Sub ReliierTablesAccess(ancienChemin As String, nouveauChemin As String)
    Dim db As DAO.Database
    Dim tdf As DAO.TableDef

    Set db = CurrentDb
    For Each tdf In db.TableDefs
        ' Ne reconnecter que les liens vers le back-end concerné
        If Len(tdf.Connect) > 0 And InStr(tdf.Connect, ancienChemin) > 0 Then
            tdf.Connect = ";DATABASE=" & nouveauChemin
            tdf.RefreshLink
        End If
    Next tdf

    Set tdf = Nothing: Set db = Nothing
End Sub
```

Point essentiel : on **filtre** les liens à reconnecter plutôt que de tous les repointer aveuglément. Une application peut avoir des tables liées à plusieurs sources (un back-end Access *et* un serveur ODBC) ; on ne reconnecte que celles qui visent le back-end déplacé, en testant le contenu de `Connect`. `RefreshLink` est préférable à une suppression-recréation : il préserve l'objet `TableDef` et ses propriétés, et il est plus efficace.

## Vérifier et détecter un lien rompu

Si le back-end a été déplacé ou est inaccessible, la première tentative d'accès à une table liée échoue avec une erreur (typiquement « fichier introuvable » ; voir l'[annexe C](/annexes/c-codes-erreur-access-dao-ado.md) pour les codes). Plutôt que de laisser cette erreur surgir devant l'utilisateur, on **teste** la validité des liens au démarrage et on réagit.

```vba
Public Function LiensValides(nomTableTest As String) As Boolean
    On Error Resume Next
    Dim n As Long
    n = DCount("*", nomTableTest)        ' tente d'accéder à une table liée
    LiensValides = (Err.Number = 0)
    On Error GoTo 0
End Function
```

Si les liens sont invalides, on tente une reliaison vers le chemin attendu ; et si cela échoue encore (back-end vraiment introuvable), on invite l'utilisateur à le localiser au moyen d'une boîte de dialogue de fichier (`FileDialog`), puis on relie avec le chemin choisi.

## Le patron de démarrage

En assemblant ces éléments, on obtient le scénario robuste exécuté à l'ouverture d'une application partagée :

1. **déterminer le chemin attendu du back-end** — depuis une table de configuration, une TempVar, un fichier de paramètres, le registre, ou un emplacement connu ;
2. **tester un lien** vers ce back-end ;
3. si le lien est rompu ou pointe ailleurs, **relier** les tables vers le chemin attendu (`Connect` + `RefreshLink`) ;
4. si la reliaison échoue, **demander à l'utilisateur** de localiser le back-end, puis relier.

Ce mécanisme, qui rend l'application portable d'un environnement à l'autre, constitue le cœur de la reliaison automatique traitée du point de vue déploiement à la [section 21.4](/21-deploiement-distribution/04-reliaison-automatique-tables.md), en lien avec les chemins UNC ([section 21.5](/21-deploiement-distribution/05-deploiement-reseau-chemins-unc.md)).

Sur le plan des performances, il est inutile de reconnecter toutes les tables à chaque démarrage si le chemin n'a pas changé : on **teste d'abord** un lien, et l'on ne relie que si nécessaire. Le maintien d'une connexion persistante au back-end est par ailleurs un levier d'optimisation réseau (cf. [section 18.8](/18-optimisation-performance/08-optimisation-reseau.md)).

## Supprimer un lien

La suppression d'un lien se fait par son nom et n'affecte en rien les données du back-end :

```vba
db.TableDefs.Delete "Clients"
```

## Points de vigilance

- **Chaîne `Connect` d'un back-end Access** : `;DATABASE=<chemin>`, point-virgule initial inclus, sans préfixe de fournisseur ; `ODBC;...` pour une source externe.
- **`RefreshLink` plutôt que supprimer-recréer** : préserve l'objet et ses propriétés, plus efficace.
- **Filtrer les liens à reconnecter** : ne repointer que ceux qui visent le back-end déplacé, jamais aveuglément tous les liens.
- **Tester avant de relier** : ne reconnecter qu'en cas de besoin, pour ne pas pénaliser le démarrage.
- **Détecter le lien rompu et prévoir un recours** : interception de l'erreur, reliaison automatique, puis invite à localiser le back-end.
- **`dbAttachSavePWD` = mot de passe stocké dans le front-end** : commodité contre sécurité, à arbitrer (chapitre 20).

## En résumé

Lier une table par code consiste à créer un `TableDef` dont on renseigne `SourceTableName` (la table dans la base source) et `Connect` (la chaîne de connexion), avant de l'ajouter à la collection `TableDefs`. Pour un back-end Access, `Connect` vaut `;DATABASE=<chemin>` ; pour une source ODBC, elle commence par `ODBC;` et décrit pilote et serveur. L'opération clé du déploiement est la **reliaison** : mettre à jour `Connect` puis appeler `RefreshLink`, en ne repointant que les liens visant le back-end concerné. Un patron de démarrage robuste détermine le chemin attendu, teste un lien, relie si nécessaire, et invite l'utilisateur à localiser le back-end en dernier recours. C'est ce mécanisme qui rend une application front-end / back-end portable et résiliente face aux changements de chemin.

La dernière section du chapitre prend du recul sur l'ensemble : les [limites du moteur ACE en environnement concurrent](10-limites-moteur-ace.md).

⏭️ [15.10. Limites du moteur ACE en environnement concurrent](/15-multi-utilisateurs/10-limites-moteur-ace.md)
