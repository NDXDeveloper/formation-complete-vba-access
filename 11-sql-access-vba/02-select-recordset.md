🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 11.2. Requêtes SELECT et ouverture de Recordset

La section précédente s'est conclue sur une distinction nette : exécuter une *action* et *lire* des données sont deux opérations différentes. Un `SELECT` ne s'« exécute » pas comme un `UPDATE` — il renvoie un **ensemble de résultats** que l'on doit ouvrir sous forme de **recordset** pour en parcourir les enregistrements. Le recordset devient alors l'image en mémoire du résultat de la requête.

Les mécanismes de parcours et de lecture d'un recordset ont déjà été détaillés au chapitre 9 (DAO) et au chapitre 10 (ADO). Cette section ne les répète pas : elle se place résolument du **point de vue du SQL**. Comment ouvre-t-on un recordset à partir d'un `SELECT` ? Quelles bonnes pratiques côté requête conditionnent la performance et la fiabilité ? C'est ce que nous examinons ici.

---

## Lire un SELECT : ouvrir un recordset

Le principe est constant, quelle que soit la technologie : on écrit une instruction `SELECT`, on **ouvre un recordset** sur cette instruction, puis on parcourt les enregistrements obtenus. Les deux voies — DAO et ADO — diffèrent dans la syntaxe d'ouverture, mais convergent dans la logique.

### En DAO : `OpenRecordset`

Avec DAO, on utilise la méthode `OpenRecordset` d'un objet `Database`. Sa source peut être une instruction SQL, un nom de table ou un nom de requête sauvegardée :

```vba
Dim db As DAO.Database
Dim rs As DAO.Recordset

Set db = CurrentDb
Set rs = db.OpenRecordset( _
    "SELECT NomClient, Ville FROM Clients WHERE Region = 'Nord' ORDER BY NomClient;", _
    dbOpenSnapshot)

Do While Not rs.EOF
    Debug.Print rs!NomClient, rs!Ville
    rs.MoveNext
Loop

rs.Close: Set rs = Nothing
Set db = Nothing
```

Le type de recordset (deuxième argument) influe sur le comportement et la performance : pour une **lecture seule en un seul passage**, `dbOpenSnapshot` ou `dbOpenForwardOnly` sont les plus rapides (le choix des types est détaillé au chapitre 9.3).

### En ADO : `Recordset.Open`

Avec ADO, on emploie la méthode `Open` du `Recordset`, en s'appuyant sur une connexion (par exemple `CurrentProject.Connection`) :

```vba
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset

rs.Open "SELECT NomClient, Ville FROM Clients WHERE Region = 'Nord' ORDER BY NomClient;", _
        CurrentProject.Connection, adOpenForwardOnly, adLockReadOnly, adCmdText

Do While Not rs.EOF
    Debug.Print rs!NomClient, rs!Ville
    rs.MoveNext
Loop

rs.Close: Set rs = Nothing
```

Pour une simple lecture, la combinaison *firehose* (`adOpenForwardOnly` + `adLockReadOnly`) est la plus rapide (section 10.4).

> 📌 L'essentiel de cette section concernant le SQL, les deux exemples ci-dessus suffisent à fixer la mécanique d'ouverture. Pour tout le détail du parcours, de la lecture des champs, de la gestion des `Null` et des types de recordset, reportez-vous aux chapitres 9 et 10.

---

## Le SELECT fait le travail : filtrer et trier à la source

Voici le principe le plus important de cette section, et l'un des plus structurants de tout le développement Access : **laissez la requête faire le travail**. Le filtrage (`WHERE`) et le tri (`ORDER BY`) doivent figurer **dans le `SELECT`**, et non être réalisés dans une boucle VBA après coup.

La différence de performance est considérable. Filtrer dans la clause `WHERE` permet au moteur d'exploiter les **index** et de ne renvoyer que les lignes utiles ; trier via `ORDER BY` confie l'opération au moteur, optimisé pour cela. À l'inverse, charger une table entière pour la filtrer ou la trier ensuite enregistrement par enregistrement constitue un **anti-patron** classique : on transfère inutilement des milliers de lignes, pour en écarter la plupart côté client.

```vba
' À PRIVILÉGIER : le moteur filtre et trie
Set rs = db.OpenRecordset( _
    "SELECT * FROM Commandes WHERE Montant > 1000 ORDER BY DateCommande DESC;", _
    dbOpenSnapshot)

' À ÉVITER : tout charger puis filtrer en VBA
Set rs = db.OpenRecordset("SELECT * FROM Commandes;", dbOpenSnapshot)
Do While Not rs.EOF
    If rs!Montant > 1000 Then ' ... ' tri et filtre manuels : lent et fastidieux
    rs.MoveNext
Loop
```

Cette règle se prolonge dans les considérations de performance du chapitre 18.

---

## Ne lire que ce qu'il faut : colonnes, `TOP`, `DISTINCT`

Dans le même esprit, on restreint le résultat au strict nécessaire.

**Sélectionner les colonnes utiles** plutôt que `SELECT *` réduit le volume transféré et clarifie l'intention. Le `*` reste pratique pour un usage ponctuel, mais nommer les colonnes est préférable dès qu'on n'a besoin que de certaines d'entre elles.

La clause **`TOP n`** limite le nombre de lignes renvoyées — utile pour un aperçu, un classement, ou les *n* premiers résultats : `SELECT TOP 10 NomClient FROM Clients ORDER BY ChiffreAffaires DESC;`. (Une particularité du dialecte Jet, abordée au chapitre 11.8, mérite d'être notée : `TOP` renvoie aussi les ex æquo à la position limite.)

La clause **`DISTINCT`** élimine les doublons : `SELECT DISTINCT Ville FROM Clients;`. Elle est précieuse pour obtenir des valeurs uniques, par exemple pour alimenter une liste déroulante.

---

## Colonnes calculées et agrégats : pensez aux alias

Lorsqu'un `SELECT` contient une **expression calculée** ou une **fonction d'agrégat** (`COUNT`, `SUM`, `AVG`…), il faut lui attribuer un **alias** avec le mot-clé `AS` — faute de quoi sa lecture devient hasardeuse.

```vba
Set rs = db.OpenRecordset( _
    "SELECT Prix * Quantite AS Total FROM LignesCommande;", dbOpenSnapshot)
Debug.Print rs!Total   ' lisible grâce à l'alias
```

La raison est concrète : **sans alias, le moteur Jet génère un nom automatique** du type `Expr1000`, `Expr1001`, etc., imprévisible et peu maniable. En nommant explicitement la colonne (`AS Total`), on garantit un accès fiable par son nom. La règle vaut tout autant pour les agrégats :

```vba
Set rs = db.OpenRecordset("SELECT COUNT(*) AS NbClients FROM Clients;", dbOpenSnapshot)
Debug.Print rs!NbClients
```

**Aliasez systématiquement** toute colonne calculée ou agrégée que vous comptez relire par son nom.

---

## Lire une valeur unique (scalaire)

Lorsqu'un `SELECT` ne renvoie qu'**une seule valeur** — un comptage, un maximum, une recherche ponctuelle —, on ouvre le recordset et l'on lit simplement son premier champ :

```vba
Set rs = db.OpenRecordset("SELECT MAX(DateCommande) AS DerniereDate FROM Commandes;", dbOpenSnapshot)
If Not rs.EOF Then Debug.Print rs!DerniereDate
```

Mais pour ce type de besoin simple, il existe souvent plus direct encore : les **fonctions de domaine** (`DLookup`, `DCount`, `DSum`, `DMax`…), qui retournent une valeur unique sans ouvrir explicitement de recordset. Elles font l'objet de la section 11.10, et constituent fréquemment le moyen le plus concis d'obtenir une valeur isolée.

---

## SELECT avec critères variables

Très souvent, le `SELECT` doit intégrer des **critères variables** — une ville choisie à l'écran, une période, un identifiant. Deux approches s'offrent alors, déjà esquissées et développées plus loin dans ce chapitre.

La première consiste à **construire la chaîne SQL dynamiquement**, en y insérant les valeurs. C'est souple, mais cela exige une grande rigueur de formatage (dates, nombres, apostrophes) et de sécurité — sujets des sections 11.6 et 11.7.

La seconde, plus sûre, repose sur le **paramétrage** : requêtes paramétrées via l'objet `Command` d'ADO (section 10.6), via une `QueryDef` paramétrée en DAO (chapitre 12.3), ou via le SQL paramétré (section 11.5). C'est l'approche à privilégier dès que des valeurs d'origine externe entrent dans la requête.

---

## Recordset modifiable ou en lecture seule ?

Un recordset ouvert sur un `SELECT` n'est pas nécessairement modifiable. Sa **mise à jour** dépend de la nature de la requête : un `SELECT` sur **une seule table** est généralement modifiable, tandis qu'une requête comportant des **agrégats**, un `DISTINCT`, une `UNION` ou certaines jointures ne l'est pas. Pour une simple lecture, peu importe — un snapshot suffit. Mais si vous comptez **modifier les données via le recordset**, assurez-vous que le `SELECT` est modifiable et que le type/verrouillage le permet (chapitres 9.3 et 10.7).

---

## SQL en ligne ou requête sauvegardée ?

On peut ouvrir un recordset sur une instruction SQL **en ligne** (comme dans tous les exemples ci-dessus) ou sur le **nom d'une requête sauvegardée** dans la base. Le choix relève d'un compromis.

Le **SQL en ligne** est souple et permet une construction dynamique. La **requête sauvegardée**, elle, est **précompilée** — Access conserve un plan d'exécution optimisé —, ce qui peut la rendre plus rapide pour un usage répété, tout en sortant le SQL du code (meilleure lisibilité et maintenance). On l'invoque simplement par son nom : `db.OpenRecordset("MaRequeteSauvegardee")`. La manipulation des requêtes sauvegardées par code est traitée au chapitre 12.1.

---

## Gérer un résultat vide

Comme toujours, on ne présume jamais qu'une requête a renvoyé des lignes. Avant de lire, on teste que le recordset n'est pas vide — `If Not (rs.BOF And rs.EOF) Then …` —, test détaillé pour DAO comme pour ADO aux chapitres 9 et 10. Un `SELECT` parfaitement valide peut légitimement ne renvoyer aucun enregistrement.

---

## En résumé

Un `SELECT` ne s'exécute pas : on **ouvre un recordset** sur son résultat, via `OpenRecordset` en DAO ou `Recordset.Open` en ADO, pour en parcourir les enregistrements. Le principe directeur est de **laisser la requête travailler** — filtrer (`WHERE`) et trier (`ORDER BY`) à la source plutôt qu'en VBA — et de ne lire que le nécessaire (colonnes ciblées, `TOP`, `DISTINCT`). Toute colonne calculée ou agrégée doit recevoir un **alias** (`AS`) pour être relue de façon fiable, sous peine de noms automatiques (`Expr1000`). Pour une valeur unique, les fonctions de domaine (section 11.10) sont souvent plus directes. Enfin, la modifiabilité du recordset dépend de la requête, et l'on choisit entre SQL en ligne (souple) et requête sauvegardée (précompilée) selon le contexte.

Nous savons désormais lire des données via un `SELECT`. L'autre versant — **modifier** les données — passe par les requêtes action. La section suivante détaille `INSERT`, `UPDATE` et `DELETE` : leur syntaxe, leurs usages et leurs pièges.


⏭️ [11.3. Requêtes action : INSERT, UPDATE, DELETE](/11-sql-access-vba/03-requetes-action-insert-update-delete.md)
