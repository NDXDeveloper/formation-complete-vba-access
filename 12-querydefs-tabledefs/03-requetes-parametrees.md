🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 12.3. Requêtes paramétrées — définition et exécution par code

La section 12.2 a couvert le cycle de vie d'une requête en laissant de côté un aspect essentiel : les **paramètres**. Nous y venons. Le paramétrage d'une requête a déjà été présenté au chapitre 11.5 comme la réponse à la fois à la **sécurité** (prévention des injections) et à la **correction** (valeurs typées, sans souci de formatage). Cette section en détaille la mise en œuvre côté DAO : définir une requête paramétrée et l'exécuter par code.

Rappelons la différence fondamentale avec ADO (section 10.6) : là où ADO emploie des paramètres **positionnels** (`?`), DAO utilise des paramètres **nommés**, déclarés par une clause `PARAMETERS`. C'est cette voie nommée que nous explorons ici.

---

## La clause `PARAMETERS` : déclarer des paramètres nommés

Une requête Jet paramétrée commence par une clause **`PARAMETERS`** qui déclare chaque paramètre par son **nom** et son **type**. Ces noms servent ensuite de marqueurs dans le corps de la requête :

```sql
PARAMETERS pRegion TEXT(50), pMontantMin CURRENCY;
SELECT * FROM Clients WHERE Region = pRegion AND Solde >= pMontantMin;
```

Les types employés sont les **noms de types Jet SQL** vus à la section 11.4 (`TEXT(n)`, `SHORT`, `LONG`, `DOUBLE`, `CURRENCY`, `DATETIME`, `YESNO`…). Le paramètre `pRegion` est ici déclaré comme texte de 50 caractères, `pMontantMin` comme monétaire.

---

## Pourquoi déclarer les types ? Ce n'est pas cosmétique

On pourrait croire la clause `PARAMETERS` facultative — et de fait, Jet tolère les **paramètres implicites** : tout identifiant qu'il ne reconnaît pas dans une requête est traité comme un paramètre. Écrire `WHERE Region = [Quelle région ?]` transforme `[Quelle région ?]` en paramètre, et Access **invite l'utilisateur** à saisir une valeur via une boîte de dialogue.

Mais ces paramètres implicites sont **non typés** — traités comme du texte. Comparer une date ou un nombre via un paramètre non déclaré peut alors provoquer une **incompatibilité de type** ou un résultat erroné. La clause `PARAMETERS` n'est donc pas un ornement : en **déclarant explicitement le type**, elle garantit que le moteur compare correctement une date à une date, un nombre à un nombre. Pour un usage par code, **on déclare toujours ses paramètres explicitement**.

---

## Définir les valeurs et exécuter

Une fois la requête paramétrée en main, on affecte les valeurs via la collection **`Parameters`** de l'objet `QueryDef`, **par nom**, puis on ouvre un recordset (sélection) ou on exécute (action).

### Sur une requête sauvegardée

```vba
Dim qdf As DAO.QueryDef
Dim rs As DAO.Recordset

Set qdf = CurrentDb.QueryDefs("R_ClientsParRegion")   ' requête paramétrée sauvegardée
qdf.Parameters("pRegion") = "Nord"
qdf.Parameters("pMontantMin") = 1000

Set rs = qdf.OpenRecordset(dbOpenSnapshot)
```

Pour une requête action paramétrée, on remplace `OpenRecordset` par `Execute` :

```vba
qdf.Parameters("pID") = 42
qdf.Parameters("pStatut") = "Traité"
qdf.Execute dbFailOnError
```

### Sur une requête temporaire

C'est le patron canonique présenté au chapitre 11.5 : une `QueryDef` **temporaire** (nom vide), paramétrée, pour une exécution ponctuelle sans encombrer la base d'une requête sauvegardée.

```vba
Dim qdf As DAO.QueryDef
Set qdf = CurrentDb.CreateQueryDef("", _
    "PARAMETERS pRegion TEXT(50); " & _
    "SELECT * FROM Clients WHERE Region = pRegion;")

qdf.Parameters("pRegion") = Me.txtRegion           ' valeur fournie séparément

Dim rs As DAO.Recordset
Set rs = qdf.OpenRecordset(dbOpenSnapshot)
```

---

## Le bénéfice clé : aucun formatage

Voici la résolution concrète du second fil rouge du chapitre 11. Parce qu'on affecte aux paramètres des **valeurs VBA natives**, on échappe **entièrement** aux pièges de formatage de la section 11.7 — dièses autour des dates, point décimal, doublement des apostrophes. Le moteur reçoit une valeur typée et la traite correctement, quels que soient les paramètres régionaux du poste :

```vba
qdf.Parameters("pDate") = Date              ' date native, aucun dièse ni format à gérer
qdf.Parameters("pMontant") = 1234.56        ' décimale native, aucun souci de séparateur
qdf.Parameters("pNom") = "O'Connor"         ' apostrophe gérée, aucune à doubler
```

C'est exactement le bénéfice qu'apportent les paramètres ADO (section 10.6), et la raison pour laquelle le paramétrage est *la* solution recommandée tant pour la sécurité que pour le formatage.

---

## Réutiliser une requête paramétrée

Comme l'objet `Command` d'ADO avec sa propriété `Prepared`, une `QueryDef` paramétrée se **réutilise** : on ne change que les valeurs entre deux exécutions. Pour une requête sauvegardée (donc précompilée), c'est particulièrement efficace dans une boucle :

```vba
Dim qdf As DAO.QueryDef
Set qdf = CurrentDb.QueryDefs("R_MajClient")   ' requête action paramétrée

Dim i As Long
For i = 1 To 100
    qdf.Parameters("pID") = i
    qdf.Parameters("pStatut") = "Traité"
    qdf.Execute dbFailOnError
Next i
```

On définit la requête une fois, et seules les valeurs varient — un gain de performance et de clarté.

---

## Éviter l'invite de paramètre

Un symptôme classique trahit un paramètre non résolu : la boîte de dialogue **« Entrez une valeur de paramètre »** surgit à l'exécution. Elle apparaît lorsqu'un paramètre attendu n'a reçu aucune valeur — par exemple parce qu'il référence un contrôle de formulaire fermé, ou qu'on a oublié de l'affecter.

**Définir les paramètres par code** (`qdf.Parameters(...) = …`) supprime cette invite. Alternativement, une requête sauvegardée peut référencer directement un contrôle de formulaire dans son critère — `Forms!frmRecherche!txtRegion` —, qu'Access résout automatiquement lorsque le formulaire est ouvert (mécanisme évoqué au chapitre 11.5). Mais pour une exécution purement programmatique, l'affectation explicite reste la voie la plus sûre.

---

## La collection Parameters

La collection `Parameters` se manipule comme les autres collections DAO (chapitre 3.7). On y accède par **nom** ou par **index**, et chaque objet `Parameter` expose `Name`, `Type` et `Value` :

```vba
Dim prm As DAO.Parameter
For Each prm In qdf.Parameters
    Debug.Print prm.Name & " (type " & prm.Type & ")"
Next prm
```

Ce parcours est utile pour du **code générique** qui affecte dynamiquement les paramètres d'une requête sans en connaître la liste à l'avance.

---

## DAO nommé, ADO positionnel, et le cas Pass-Through

Deux rappels pour situer cette approche. D'abord, la **différence de paradigme** : DAO emploie des paramètres **nommés**, dont l'**ordre n'importe pas** (on affecte par nom), tandis qu'ADO emploie des paramètres **positionnels** (`?`), liés par leur **ordre** (section 10.6). Ensuite, une **limite** : les requêtes **Pass-Through** n'acceptent **pas** la clause `PARAMETERS` ni les paramètres Jet (section 11.9). Le paramétrage DAO décrit ici concerne les requêtes natives du moteur ACE ; pour une Pass-Through, on recourt à une procédure stockée ou à la construction de SQL, comme vu précédemment.

---

## En résumé

Une requête paramétrée DAO déclare ses paramètres **nommés et typés** via une clause **`PARAMETERS`**, dont la déclaration de type n'est pas cosmétique : elle prévient les incompatibilités et les invites de paramètre. On affecte les valeurs par la collection **`Parameters`** (`qdf.Parameters("nom") = valeur`), puis on **ouvre un recordset** (sélection) ou on **exécute** (action) — sur une requête **sauvegardée** comme sur une **temporaire** (`CreateQueryDef("")`, le patron du chapitre 11.5). Le bénéfice majeur est l'**absence totale de formatage** : les valeurs natives résolvent les pièges de la section 11.7, à l'instar des paramètres ADO. La requête se **réutilise** en ne changeant que les valeurs, et l'affectation par code **évite l'invite de paramètre**. Enfin, DAO est **nommé** (ordre indifférent) là où ADO est **positionnel**, et les requêtes Pass-Through ne prennent pas en charge ce mécanisme.

Nous avons fait le tour des requêtes — accès, création, modification, suppression, paramétrage. Le chapitre bascule maintenant vers l'autre grand objet structurel : les **tables**. La section suivante introduit la collection `TableDefs` et l'inspection de la structure des tables.


⏭️ [12.4. Collection TableDefs — inspection de la structure des tables](/12-querydefs-tabledefs/04-collection-tabledefs.md)
