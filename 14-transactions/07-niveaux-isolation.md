🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 14.7. Niveaux d'isolation (ODBC / serveur lié) — limites du moteur ACE natif

L'isolation est le « I » de l'acronyme ACID introduit à la [section 14.1](01-concept-transaction.md). Elle gouverne la manière dont les transactions concurrentes se perçoivent — ou s'ignorent — les unes les autres. C'est aussi le domaine où le moteur ACE natif montre ses limites les plus nettes, et où la frontière entre une base Access autonome et une architecture connectée à un serveur devient déterminante. Cette dernière section du chapitre explique ce qu'est un niveau d'isolation, pourquoi le moteur ACE n'en propose pas de paramétrable, et comment l'isolation se gère réellement lorsqu'un serveur entre en jeu.

## Qu'est-ce qu'un niveau d'isolation ?

Un niveau d'isolation définit dans quelle mesure les modifications d'une transaction sont isolées des autres transactions qui s'exécutent en même temps. On le décrit généralement par les **effets de concurrence** qu'il autorise ou interdit, c'est-à-dire par le comportement des lectures vis-à-vis de données en cours de modification par d'autres transactions.

Trois phénomènes classiques servent de repères :

- **la lecture sale** (*dirty read*) : lire une donnée modifiée par une autre transaction non encore validée — donnée qui pourrait être annulée ensuite ;
- **la lecture non reproductible** (*non-repeatable read*) : relire le même enregistrement au sein d'une transaction et obtenir une valeur différente, car une autre transaction l'a modifié et validé entre-temps ;
- **la lecture fantôme** (*phantom read*) : ré-exécuter la même requête et voir apparaître ou disparaître des lignes, car une autre transaction en a inséré ou supprimé.

À partir de ces phénomènes, la norme SQL définit quatre niveaux d'isolation croissants : **Read Uncommitted** (autorise les lectures sales), **Read Committed** (interdit les lectures sales), **Repeatable Read** (interdit en plus les lectures non reproductibles) et **Serializable** (interdit aussi les lectures fantômes, les transactions se comportant comme si elles s'exécutaient l'une après l'autre).

Le choix d'un niveau est un arbitrage. Concrètement, les niveaux d'isolation déterminent quels verrous sont posés lors des lectures et combien de temps ils sont conservés, et si une lecture portant sur des lignes modifiées par une autre transaction doit attendre, lire la version validée, ou lire la donnée non validée. Un niveau bas augmente la concurrence — davantage d'utilisateurs accèdent aux données simultanément — mais multiplie les effets indésirables ; un niveau élevé réduit ces effets mais consomme plus de ressources et augmente les blocages. À noter qu'indépendamment du niveau, une transaction obtient toujours un verrou exclusif sur les données qu'elle modifie, conservé jusqu'à sa conclusion. Le niveau d'isolation concerne donc avant tout le comportement des **lectures**.

## Le moteur ACE natif n'expose pas de niveaux d'isolation

C'est le point central de cette section : **le moteur ACE natif ne permet pas de choisir un niveau d'isolation paramétrable**. Il n'existe pas, en Jet SQL, d'équivalent de l'instruction `SET TRANSACTION ISOLATION LEVEL` des serveurs : on ne peut pas demander à une base Access native de fonctionner en Read Committed plutôt qu'en Serializable.

La gestion de la concurrence du moteur ACE repose non pas sur des niveaux d'isolation, mais sur le **verrouillage**. L'isolation qu'offre une transaction ACE est ainsi une conséquence de sa stratégie de verrouillage (optimiste ou pessimiste, verrouillage de page ou d'enregistrement) plutôt qu'un paramètre que l'on règle. En pratique, pour une application travaillant sur des tables Access natives, on ne raisonne donc pas en termes de niveaux d'isolation, mais en termes de :

- **stratégie de verrouillage** ([section 15.3](/15-multi-utilisateurs/03-strategies-verrouillage.md)) ;
- **propriété `RecordLocks` des formulaires** ([section 15.6](/15-multi-utilisateurs/06-recordlocks-formulaires.md)) ;
- **détection et résolution des conflits de mise à jour** ([section 14.5](05-conflits-mise-a-jour.md)).

Ce sont ces leviers — et non un niveau d'isolation — qui définissent le comportement concurrent d'une base ACE autonome. Il s'agit là d'une limite structurelle du moteur, à laquelle on s'adapte par la conception plutôt que par un réglage.

## L'isolation entre en jeu avec un serveur lié

La notion de niveau d'isolation redevient pertinente — et contrôlable — dès lors qu'Access n'est plus qu'un **frontal** connecté à un serveur de base de données via ODBC (tables liées ou requêtes pass-through). Par défaut, la plupart des requêtes s'exécutent localement dans le moteur ACE ; mais lorsqu'une opération est effectivement traitée par le serveur, c'est **le serveur et la couche ODBC qui gouvernent l'isolation**, et non le moteur ACE.

Deux conséquences en découlent. D'une part, le niveau d'isolation applicable est celui du serveur : pour SQL Server, par exemple, le niveau par défaut est **Read Committed**. D'autre part, le moteur ACE n'a pas la maîtrise de ce mécanisme : c'est d'ailleurs la raison pour laquelle, on l'a vu en [section 14.5](05-conflits-mise-a-jour.md), la propriété `LockEdits` d'un recordset DAO est toujours optimiste pour une source ODBC — le moteur ACE ne contrôle pas le verrouillage des serveurs externes.

## Contrôler l'isolation côté serveur

Plusieurs leviers permettent d'agir sur l'isolation lorsqu'un serveur est dans la boucle.

**Les requêtes pass-through.** Une requête pass-through est une instruction écrite dans la syntaxe du serveur (T-SQL pour SQL Server) et envoyée telle quelle au serveur via une chaîne de connexion ODBC. C'est le moyen d'utiliser des commandes propres au serveur — y compris celles qui touchent à l'isolation. On peut ainsi soit positionner le niveau de la session avec `SET TRANSACTION ISOLATION LEVEL`, soit, plus finement, appliquer un indicateur de table à une requête précise (par exemple `WITH (NOLOCK)` pour une lecture non verrouillée). L'approche par indicateur de table a l'avantage de ne concerner que la requête visée, sans modifier le comportement des autres instructions. Les requêtes pass-through sont traitées en détail à la [section 11.9](/11-sql-access-vba/09-requetes-pass-through.md).

```vba
Public Sub LireSansVerrou()
    Dim db As DAO.Database
    Dim qdf As DAO.QueryDef
    Dim rs As DAO.Recordset

    Set db = CurrentDb
    Set qdf = db.CreateQueryDef("")        ' requête pass-through temporaire
    qdf.Connect = "ODBC;DRIVER={ODBC Driver 17 for SQL Server};" & _
                  "SERVER=MonServeur;DATABASE=MaBase;Trusted_Connection=Yes;"
    qdf.ReturnsRecords = True

    ' Syntaxe du SERVEUR (T-SQL), pas Jet SQL.
    ' WITH (NOLOCK) : lecture non verrouillée pour CETTE requête uniquement.
    qdf.SQL = "SELECT NumCommande, DateCommande FROM dbo.Commandes WITH (NOLOCK);"

    Set rs = qdf.OpenRecordset(dbOpenSnapshot)
    ' ... exploitation du recordset ...
    rs.Close

    Set rs = Nothing
    Set qdf = Nothing
    Set db = Nothing
End Sub
```

**La propriété `IsolationLevel` d'ADO.** Dans une architecture ADO connectée à un fournisseur serveur, on peut fixer le niveau d'isolation de la connexion avant d'ouvrir une transaction, via la propriété `Connection.IsolationLevel`. Les constantes correspondent aux niveaux normalisés : `adXactReadUncommitted`, `adXactReadCommitted`, `adXactRepeatableRead`, `adXactSerializable`.

```vba
Dim cn As ADODB.Connection
Set cn = New ADODB.Connection
cn.Open "Provider=MSOLEDBSQL;Server=MonServeur;Database=MaBase;Trusted_Connection=Yes;"

cn.IsolationLevel = adXactSerializable     ' à définir AVANT BeginTrans
cn.BeginTrans
' ... opérations ...
cn.CommitTrans
cn.Close
```

Il faut toutefois garder à l'esprit que l'effet réel de cette propriété **dépend du fournisseur** : elle prend tout son sens face à un fournisseur serveur capable d'honorer ces niveaux, alors que sur une base ACE native, où l'isolation paramétrable n'existe pas, elle reste sans portée véritable.

**Au niveau ODBC.** Plus bas dans la pile, l'isolation peut être fixée par l'attribut de connexion ODBC dédié. Si la source ne supporte pas le niveau demandé, le pilote ou la source peut appliquer un niveau plus élevé. Ce réglage de bas niveau est rarement manipulé directement depuis VBA, mais il explique le comportement observé : c'est bien la couche ODBC/serveur, et non ACE, qui tranche.

## Points de vigilance

- **Aucun réglage d'isolation pour ACE natif.** Sur une base Access autonome, on conçoit la concurrence avec le verrouillage et la gestion des conflits, pas avec des niveaux d'isolation.
- **`SET ... ISOLATION LEVEL` suivi d'un `SELECT` dans une même pass-through.** Combiner une instruction `SET` (qui ne renvoie aucune ligne) et un `SELECT` dans la même requête pass-through peut conduire Access à considérer que la requête ne retourne rien. Préférer l'indicateur de table par requête, ou séparer les instructions.
- **Persistance de la session.** Un niveau posé par `SET TRANSACTION ISOLATION LEVEL` vaut pour la session ; si une autre pass-through utilise une connexion distincte, elle n'en hérite pas. Les indicateurs de table, eux, sont autonomes par requête.
- **`NOLOCK` / Read Uncommitted = lectures sales.** Lire sans verrou améliore la fluidité mais expose à des données non validées, susceptibles d'être annulées. À réserver aux lectures où cette approximation est acceptable (statistiques, affichage indicatif), jamais pour une décision critique.
- **Isolation élevée = blocages accrus.** Plus le niveau est strict (jusqu'à Serializable), plus l'intégrité est protégée, mais plus les transactions risquent de s'attendre mutuellement. Le bon niveau résulte d'un équilibre entre intégrité et concurrence.
- **`IsolationLevel` ADO dépend du fournisseur.** Sans portée sur une base ACE native ; significatif face à un serveur.

## En résumé

Un niveau d'isolation décrit dans quelle mesure une transaction est protégée des effets des autres (lectures sales, non reproductibles, fantômes), avec un arbitrage permanent entre intégrité et concurrence. Le moteur ACE natif **n'expose aucun niveau d'isolation paramétrable** : sa concurrence repose sur le verrouillage, et l'on conçoit une base Access autonome avec les stratégies de verrouillage (15.3), `RecordLocks` (15.6) et la gestion des conflits (14.5), non avec un réglage d'isolation. La notion redevient pertinente lorsqu'Access sert de frontal à un serveur via ODBC : l'isolation est alors gouvernée par le serveur (Read Committed par défaut sur SQL Server) et se contrôle côté serveur, par requête pass-through (`SET TRANSACTION ISOLATION LEVEL` ou indicateur de table comme `WITH (NOLOCK)`), par la propriété `IsolationLevel` d'ADO, ou via la couche ODBC — l'effet dépendant toujours du fournisseur.

---

## Conclusion du chapitre 14

Ce chapitre a parcouru l'ensemble des moyens de garantir la fiabilité et la cohérence des données en VBA Access. Il est parti du **concept** de transaction et du principe « tout ou rien » ([14.1](01-concept-transaction.md)), puis a détaillé sa mise en œuvre concrète en [DAO](02-transactions-dao.md) et en [ADO](03-transactions-ado.md), avant d'aborder le cas particulier des [transactions imbriquées](04-transactions-imbriquees.md). Il a ensuite traité les [conflits de mise à jour](05-conflits-mise-a-jour.md) en environnement partagé, les règles d'[intégrité référentielle](06-integrite-referentielle.md) qui protègent les liens entre tables, et enfin les limites du moteur ACE en matière d'[isolation](07-niveaux-isolation.md).

Deux idées structurent l'ensemble. D'abord, **transaction et gestion d'erreur sont indissociables** : une transaction n'a de valeur que couplée à un mécanisme qui déclenche l'annulation au bon moment. Ensuite, **le moteur ACE est un moteur fichier, pas un serveur** : il offre l'atomicité et l'intégrité référentielle, mais sa gestion de la concurrence repose sur le verrouillage plutôt que sur des niveaux d'isolation configurables — limite qui motive souvent, à terme, une architecture connectée à un serveur.

Cette dernière dimension — la concurrence et le partage — est précisément l'objet du [chapitre 15, consacré au multi-utilisateurs et au verrouillage](/15-multi-utilisateurs/README.md), qui prolonge et approfondit les questions abordées ici.

⏭️ [15. Multi-utilisateurs et verrouillage](/15-multi-utilisateurs/README.md)
