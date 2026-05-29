🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 11.5. SQL paramétré — prévention des injections SQL

Toutes les requêtes des sections précédentes manipulaient des valeurs soit fixes, soit insérées directement dans la chaîne SQL par concaténation. Or c'est précisément cette **insertion directe** qui devient dangereuse dès que les valeurs proviennent de l'extérieur — d'un formulaire, d'une saisie utilisateur, d'une source non maîtrisée. Nous abordons ici l'un des deux fils rouges annoncés dans l'introduction du chapitre : la **sécurité**, à travers le **SQL paramétré** et la prévention des **injections SQL**.

L'enjeu est double. D'une part, la sécurité : une chaîne SQL construite par concaténation peut être détournée par une entrée malveillante. D'autre part, la simple correction : la concaténation se casse aussi sur des valeurs parfaitement légitimes. Le paramétrage résout les deux problèmes d'un même geste.

---

## Le danger : l'injection SQL

Considérons une recherche apparemment anodine, construite par concaténation :

```vba
strSQL = "SELECT * FROM Utilisateurs WHERE Nom = '" & Me.txtNom & "';"
```

Avec une saisie normale (`Dupont`), la requête est correcte : `… WHERE Nom = 'Dupont'`. Mais que se passe-t-il si l'utilisateur saisit une valeur conçue pour détourner la requête ? Avec `txtNom` valant `' OR '1'='1`, la chaîne devient :

```sql
SELECT * FROM Utilisateurs WHERE Nom = '' OR '1'='1';
```

La condition `'1'='1'` étant toujours vraie, la requête renvoie **tous les utilisateurs**. Sur un écran de connexion, ce type de manipulation constitue un **contournement d'authentification**. D'autres détournements existent : élargir la portée d'un `WHERE`, extraire des données via une clause `UNION` injectée, ou simplement faire échouer la requête.

### La nuance Access : un risque réel, mais à mesurer justement

Une précision honnête s'impose pour le contexte Access. Le moteur ACE **n'exécute qu'une instruction à la fois** : il ne permet pas, contrairement à SQL Server, d'empiler plusieurs requêtes séparées par des points-virgules dans un seul `Execute`. L'injection destructrice classique du type `'; DROP TABLE …; --` est donc **largement impossible** contre le moteur ACE natif.

Cela ne rend pas Access immunisé pour autant. Le contournement d'authentification, l'élargissement de portée et l'extraction par `UNION` **restent de véritables menaces**. Et surtout, dès que l'application s'appuie sur des **tables liées SQL Server**, des **requêtes pass-through** (section 11.9) ou une **base externe** (chapitre 10.9), c'est le SGBD distant qui interprète le SQL — et là, l'éventail complet des attaques, requêtes empilées comprises, redevient possible. La règle est donc sans exception : **on paramètre, quel que soit le moteur**.

### Au-delà de la sécurité : la concaténation casse aussi

Même en écartant toute malveillance, la concaténation est fragile. Une valeur légitime contenant une **apostrophe** — un nom comme `O'Connor` — brise la requête :

```sql
SELECT * FROM Clients WHERE Nom = 'O'Connor';   -- erreur de syntaxe
```

L'apostrophe interne ferme prématurément la chaîne. À cela s'ajoutent les problèmes de formatage des dates et des nombres, traités à la section 11.7. La concaténation est donc à la fois **risquée** et **peu fiable**.

---

## La solution : les requêtes paramétrées

Le principe du paramétrage est de **séparer la structure de la requête de ses valeurs**. La structure SQL est définie une fois, avec des emplacements réservés ; les valeurs sont fournies **séparément** et ne sont **jamais interprétées comme du code SQL** — seulement comme des données.

Ce mécanisme apporte deux bénéfices indissociables. La **sécurité** d'abord : une valeur ne pouvant plus altérer la structure de la requête, l'injection est neutralisée à la racine. La **correction** ensuite : l'apostrophe d'`O'Connor`, le format d'une date, le séparateur décimal d'un nombre sont gérés automatiquement par le moteur, qui reçoit des valeurs **typées** plutôt qu'un texte à interpréter. On élimine d'un coup la faille de sécurité *et* les bugs de formatage.

---

## Paramétrer en ADO

Le paramétrage ADO a été détaillé à la section 10.6 ; rappelons-en l'essentiel. Avec le moteur ACE, on utilise des emplacements **positionnels `?`**, puis on définit les paramètres via `CreateParameter` et `Append` :

```vba
Dim cmd As ADODB.Command
Set cmd = New ADODB.Command
cmd.ActiveConnection = CurrentProject.Connection
cmd.CommandText = "SELECT * FROM Utilisateurs WHERE Nom = ?;"
cmd.CommandType = adCmdText

cmd.Parameters.Append cmd.CreateParameter("pNom", adVarWChar, adParamInput, 50, Me.txtNom)

Dim rs As ADODB.Recordset
Set rs = cmd.Execute
```

Quelle que soit la valeur de `txtNom` — y compris `' OR '1'='1` —, elle est traitée comme une simple chaîne à rechercher, et non comme du SQL.

---

## Paramétrer en DAO

DAO offre son propre mécanisme via une **`QueryDef` paramétrée**, dont la définition complète figure au chapitre 12.3. À la différence du `?` positionnel d'ADO, DAO emploie des **paramètres nommés**, déclarés par une clause `PARAMETERS`. La `QueryDef` temporaire (créée avec un nom vide `""`) est le moyen idéal de paramétrer une requête ponctuelle :

```vba
Dim qdf As DAO.QueryDef
Set qdf = CurrentDb.CreateQueryDef("", _
    "PARAMETERS pNom TEXT; SELECT * FROM Utilisateurs WHERE Nom = pNom;")

qdf.Parameters("pNom") = Me.txtNom        ' la valeur est fournie séparément

Dim rs As DAO.Recordset
Set rs = qdf.OpenRecordset(dbOpenSnapshot)
```

Le même principe vaut pour une **requête action** : on définit la `QueryDef` paramétrée, on affecte les paramètres, puis on appelle `qdf.Execute dbFailOnError` au lieu d'`OpenRecordset`.

---

## Requêtes sauvegardées et références de formulaire

Une troisième voie, sans code de paramétrage explicite, mérite d'être signalée : une **requête sauvegardée** peut référencer directement un contrôle de formulaire, par exemple `Forms![frmRecherche]![txtNom]`. Access évalue alors cette référence de manière sûre, sans concaténation textuelle. C'est une forme de paramétrage implicite, pratique pour lier une requête à une interface de recherche.

---

## Avant / après : de la faille à la sécurité

Le contraste résume tout l'enjeu :

```vba
' ❌ VULNÉRABLE — la valeur est concaténée dans la structure
strSQL = "SELECT * FROM Utilisateurs WHERE Nom = '" & Me.txtNom & "';"
Set rs = CurrentDb.OpenRecordset(strSQL)

' ✅ SÛR — la valeur transite par un paramètre, hors de la structure
Set qdf = CurrentDb.CreateQueryDef("", _
    "PARAMETERS pNom TEXT; SELECT * FROM Utilisateurs WHERE Nom = pNom;")
qdf.Parameters("pNom") = Me.txtNom
Set rs = qdf.OpenRecordset(dbOpenSnapshot)
```

La seconde version est immunisée contre l'injection *et* tolère sans broncher un nom contenant une apostrophe.

---

## Ce que les paramètres ne peuvent pas faire : les identifiants

Une limite fondamentale doit être comprise : les paramètres substituent des **valeurs**, jamais des **identifiants ni des éléments de structure**. On ne peut pas paramétrer un **nom de table**, un **nom de colonne**, un **opérateur** ou le sens d'un `ORDER BY`. Ces éléments font partie de la *structure* de la requête, pas de ses données.

```sql
-- IMPOSSIBLE : un paramètre ne peut pas remplacer un nom de colonne
SELECT * FROM Clients ORDER BY ?;     -- ne fonctionne pas
```

Dès que la partie variable est un identifiant — une colonne de tri choisie à l'écran, par exemple —, le paramétrage est inopérant et l'on doit construire la chaîne. Mais y insérer directement une saisie utilisateur recréerait une faille. La parade est la **liste blanche** (*whitelist*) : on valide la valeur contre un **ensemble fixe d'identifiants autorisés**, sans jamais injecter le texte brut de l'utilisateur.

```vba
' L'utilisateur choisit un tri ; on ne fait JAMAIS confiance à sa saisie brute
Dim colTri As String
Select Case Me.cboTri.Value
    Case "Nom":   colTri = "NomClient"
    Case "Ville": colTri = "Ville"
    Case Else:    colTri = "IDClient"     ' valeur sûre par défaut
End Select
strSQL = "SELECT * FROM Clients ORDER BY " & colTri & ";"   ' colTri est maîtrisé
```

Ainsi, **les valeurs se paramètrent, les identifiants se valident par liste blanche**. Cette distinction est au cœur de la construction dynamique de SQL, sujet de la section 11.6.

---

## L'échappement, solution de dernier recours

S'il fallait absolument insérer une chaîne en clair — ce qu'on évite —, l'échappement minimal consiste à **doubler les apostrophes** : `Replace(valeur, "'", "''")`. C'est le strict nécessaire pour ne pas casser la requête sur un `O'Connor`. Mais cette technique est **fragile** : il suffit d'en oublier une occurrence, et elle ne règle en rien le formatage des dates ou des nombres (section 11.7). Le paramétrage reste strictement plus sûr et doit demeurer le choix par défaut ; l'échappement n'est qu'un pis-aller, jamais une stratégie.

---

## Vigilance ailleurs : fonctions de domaine et filtres

Le risque d'injection ne se limite pas aux requêtes ouvertes en recordset. Les **fonctions de domaine** (`DLookup`, `DCount`… section 11.10) reçoivent un critère sous forme de chaîne, tout aussi vulnérable s'il est construit à partir d'une saisie utilisateur. De même, affecter à `Form.Filter` ou à `RecordSource` une chaîne concaténée avec une entrée externe expose à la manipulation de la logique. La même discipline — paramétrer les valeurs, valider les identifiants — s'applique partout où une saisie externe rencontre du SQL.

---

## En résumé

Construire du SQL en y **concaténant** des valeurs externes est à la fois dangereux (injection SQL : contournement d'authentification, élargissement de portée, extraction par `UNION` — et l'arsenal complet face à un back-end externe) et peu fiable (rupture sur une apostrophe légitime). La parade est le **SQL paramétré**, qui sépare la structure des valeurs : en ADO via les emplacements `?` et `CreateParameter` (section 10.6), en DAO via une **`QueryDef` paramétrée** à paramètres nommés (chapitre 12.3). Les paramètres ne peuvent toutefois remplacer que des **valeurs**, non des **identifiants** : pour une colonne ou une table variable, on recourt à une **liste blanche** d'identifiants autorisés. L'échappement par doublement d'apostrophes n'est qu'un dernier recours, et la vigilance s'étend aux fonctions de domaine comme aux filtres de formulaire.

Les paramètres traitent élégamment les *valeurs* variables, mais certaines requêtes exigent une véritable construction dynamique — structure conditionnelle, critères combinés, identifiants variables. La section suivante en expose les bonnes pratiques et les pièges, en prolongeant la distinction valeurs/identifiants que nous venons d'établir.


⏭️ [11.6. Construction dynamique de SQL — bonnes pratiques et pièges](/11-sql-access-vba/06-construction-dynamique-sql.md)
