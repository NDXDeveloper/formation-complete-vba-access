🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 10.6. L'objet Command et les paramètres de requête

Jusqu'ici, nos commandes étaient transmises sous forme de **chaînes SQL brutes** — directement à `Connection.Execute` ou à `Recordset.Open`. Cette approche fonctionne, mais elle montre vite ses limites dès que des **valeurs variables** entrent en jeu : risques d'injection SQL, casse-tête de formatage des dates et des nombres, échappement laborieux des apostrophes. L'objet **`Command`** apporte la réponse propre à ces problèmes : il permet de **paramétrer** une requête, c'est-à-dire de séparer la structure SQL des valeurs qu'elle manipule.

C'est l'objet de prédilection pour exécuter des requêtes sécurisées, réutilisables et, le cas échéant, des procédures stockées. Cette section en présente le fonctionnement, avec une attention particulière à une spécificité d'Access qui surprend souvent : la manière dont s'écrivent les paramètres.

---

## Pourquoi l'objet Command ?

Trois bénéfices justifient le détour par un objet `Command` plutôt qu'une simple chaîne.

La **sécurité** vient en premier. Concaténer des valeurs saisies par l'utilisateur dans une chaîne SQL ouvre la porte aux **injections SQL** — l'une des failles les plus répandues. Les paramètres éliminent ce risque à la racine, car les valeurs ne sont jamais interprétées comme du code SQL. Ce sujet, central, est traité plus largement au chapitre 11.5.

La **correction** vient ensuite. Avec des paramètres, on transmet des **valeurs typées natives** : plus besoin de formater une date au format US/ISO, de gérer le séparateur décimal selon les paramètres régionaux, ou de doubler les apostrophes dans un nom comme « O'Connor ». ADO et le fournisseur s'en chargent (problèmes détaillés au chapitre 11.7).

La **réutilisation** et les **procédures stockées** complètent le tableau. Une commande peut être *préparée* une fois puis exécutée plusieurs fois avec des valeurs différentes, ce qui améliore les performances ; et c'est la voie naturelle pour appeler des procédures stockées avec arguments, notamment sur SQL Server.

---

## Anatomie de l'objet Command

L'objet `Command` s'articule autour de quelques propriétés et de deux méthodes essentielles.

La propriété **`ActiveConnection`** désigne la connexion à utiliser (un objet `Connection`, ou une chaîne de connexion). **`CommandText`** contient le texte de la commande — instruction SQL, nom de table ou nom de procédure. **`CommandType`** précise la nature de `CommandText` (`adCmdText` pour du SQL, `adCmdStoredProc` pour une procédure, `adCmdTable`…), ce qui évite à ADO de la deviner. **`Prepared`**, lorsqu'elle vaut `True`, demande au fournisseur de compiler la commande en vue d'exécutions répétées. Enfin, la collection **`Parameters`** héberge les paramètres.

Côté méthodes, **`Execute`** lance la commande (et renvoie un `Recordset` pour un `SELECT`), tandis que **`CreateParameter`** fabrique un objet `Parameter` — sans toutefois l'ajouter automatiquement à la collection, comme nous le verrons.

```vba
Dim cmd As ADODB.Command
Set cmd = New ADODB.Command
cmd.ActiveConnection = cn
cmd.CommandText = "SELECT * FROM Clients WHERE Ville = ?"
cmd.CommandType = adCmdText
```

---

## Les marqueurs de paramètres : le `?` de Jet/ACE

Voici la spécificité qui déroute le plus souvent. Avec le moteur **ACE/Jet**, les paramètres s'écrivent à l'aide d'un **point d'interrogation `?`**, et ils sont **positionnels** : c'est leur *ordre* dans la requête qui compte, pas leur nom.

```vba
cmd.CommandText = "SELECT * FROM Clients WHERE Ville = ? AND Actif = ?"
'                                                 ↑ 1er ?        ↑ 2e ?
```

Concrètement, le **premier paramètre ajouté** à la collection se lie au **premier `?`**, le deuxième au deuxième, et ainsi de suite. Le nom que l'on donne à un paramètre via `CreateParameter` est, pour ACE, **purement informatif** : il ne sert pas à la liaison. Inverser l'ordre des paramètres ou en oublier un produit donc des résultats erronés, sans nécessairement déclencher d'erreur.

> ⚠️ **Piège fréquent** : les paramètres **nommés** de la forme `@nom`, familiers aux utilisateurs de SQL Server (T-SQL), **ne fonctionnent pas** dans le SQL ad hoc adressé au moteur ACE. Pour Access, c'est le `?` positionnel qui prévaut. Les paramètres nommés relèvent du contexte des bases externes et des procédures stockées SQL Server, abordé en section 10.9.

---

## Définir les paramètres : `CreateParameter` + `Append`

La construction d'un paramètre se fait en deux temps : on le **crée** avec `CreateParameter`, puis on l'**ajoute** à la collection avec `Append`. La méthode `CreateParameter` admet la signature suivante :

```vba
cmd.CreateParameter(Nom, Type, Direction, Taille, Valeur)
```

Le **`Nom`** est informatif (pour ACE). Le **`Type`** est une constante ADO décrivant le type de donnée (`adVarWChar`, `adInteger`, `adDate`…). La **`Direction`** vaut généralement `adParamInput` (paramètre d'entrée). La **`Taille`** est **obligatoire pour les types de longueur variable** comme le texte, et ignorée pour les types de taille fixe (entier, booléen, date). La **`Valeur`** initialise le paramètre.

L'idiome canonique combine création et ajout sur une seule ligne :

```vba
cmd.Parameters.Append cmd.CreateParameter("pVille", adVarWChar, adParamInput, 50, "Lyon")
cmd.Parameters.Append cmd.CreateParameter("pActif", adBoolean, adParamInput, , True)
```

Notez l'argument `Taille` renseigné (`50`) pour le paramètre texte, mais **omis** pour le booléen — d'où la virgule isolée qui saute l'argument. Respecter cette règle évite des erreurs de troncature ou de type.

---

## Exemple : un `SELECT` paramétré

Voici une lecture paramétrée complète, de la création de la commande à l'exploitation du recordset renvoyé :

```vba
Dim cmd As ADODB.Command
Set cmd = New ADODB.Command
cmd.ActiveConnection = cn
cmd.CommandText = "SELECT * FROM Clients WHERE Ville = ? AND Actif = ?"
cmd.CommandType = adCmdText

cmd.Parameters.Append cmd.CreateParameter("pVille", adVarWChar, adParamInput, 50, "Lyon")
cmd.Parameters.Append cmd.CreateParameter("pActif", adBoolean, adParamInput, , True)

Dim rs As ADODB.Recordset
Set rs = cmd.Execute   ' renvoie un Recordset pour un SELECT

Do While Not rs.EOF
    Debug.Print rs!NomClient
    rs.MoveNext
Loop

rs.Close: Set rs = Nothing
Set cmd = Nothing
```

Aucune valeur n'est concaténée dans le SQL : la ville et le statut transitent par les paramètres, à l'abri de toute injection comme de tout souci de formatage.

---

## Exemple : une requête action paramétrée

Pour une requête action (`INSERT`, `UPDATE`, `DELETE`), on exploite le premier argument d'`Execute`, qui reçoit en sortie le **nombre de lignes affectées** :

```vba
Dim cmd As ADODB.Command
Set cmd = New ADODB.Command
cmd.ActiveConnection = cn
cmd.CommandText = "UPDATE Clients SET Ville = ? WHERE IDClient = ?"
cmd.CommandType = adCmdText

cmd.Parameters.Append cmd.CreateParameter("pVille", adVarWChar, adParamInput, 50, "Marseille")
cmd.Parameters.Append cmd.CreateParameter("pID", adInteger, adParamInput, , 42)

Dim lngLignes As Long
cmd.Execute lngLignes
Debug.Print lngLignes & " ligne(s) modifiée(s)."

Set cmd = Nothing
```

---

## Réutiliser un Command : `Prepared` et changement de valeurs

L'un des atouts majeurs de l'objet `Command` est la **réutilisation**. Plutôt que de reconstruire une commande pour chaque exécution, on la prépare une fois puis on se contente de **changer la valeur des paramètres** entre deux appels :

```vba
cmd.CommandText = "UPDATE Clients SET Ville = ? WHERE IDClient = ?"
cmd.CommandType = adCmdText
cmd.Prepared = True   ' compilation en vue d'exécutions répétées

cmd.Parameters.Append cmd.CreateParameter("pVille", adVarWChar, adParamInput, 50)
cmd.Parameters.Append cmd.CreateParameter("pID", adInteger, adParamInput)

' Première exécution
cmd.Parameters("pVille").Value = "Lyon"
cmd.Parameters("pID").Value = 12
cmd.Execute

' Deuxième exécution, mêmes structure et types, valeurs différentes
cmd.Parameters("pVille").Value = "Nantes"
cmd.Parameters("pID").Value = 27
cmd.Execute
```

La propriété **`Prepared = True`** suggère au fournisseur de compiler la commande une seule fois, ce qui accélère les exécutions successives — un gain appréciable dans une boucle d'insertion ou de mise à jour répétée.

---

## Le raccourci : passer les valeurs à `Execute`

ADO autorise une écriture plus brève, en passant directement un **tableau de valeurs** au deuxième argument d'`Execute`, sans construire la collection `Parameters` :

```vba
Set rs = cmd.Execute(, Array("Lyon", True))   ' valeurs dans l'ordre des ?
```

Ce raccourci est tentant, mais il a un coût caché : faute de paramètres explicitement définis, ADO doit parfois **interroger le fournisseur** (via `Parameters.Refresh`) pour découvrir les types attendus — un aller-retour supplémentaire, parfois capricieux avec ACE. Pour du code robuste et performant, **définissez vos paramètres explicitement** avec `CreateParameter` plutôt que de vous reposer sur cette inférence automatique.

---

## Les types de données des paramètres

Choisir la bonne constante de type est essentiel à la fiabilité. Le tableau suivant établit la correspondance entre les types de champs Access courants, les constantes ADO et les types VBA.

| Type de champ Access | Constante ADO | Type VBA | Taille requise |
|---|---|---|---|
| Texte court | `adVarWChar` | `String` | **Oui** |
| Texte long / Mémo | `adLongVarWChar` | `String` | — |
| Entier long / NuméroAuto | `adInteger` | `Long` | Non |
| Entier | `adSmallInt` | `Integer` | Non |
| Réel double | `adDouble` | `Double` | Non |
| Monétaire | `adCurrency` | `Currency` | Non |
| Date/Heure | `adDate` (ou `adDBTimeStamp`) | `Date` | Non |
| Oui/Non | `adBoolean` | `Boolean` | Non |
| Décimal | `adNumeric` | — | Précision / échelle |

Les champs texte d'Access étant stockés en Unicode, le type `adVarWChar` (texte Unicode de longueur variable) est le choix habituel, sans oublier sa **taille** obligatoire.

---

## Paramètres de sortie et procédures stockées

La propriété **`Direction`** d'un paramètre ne se limite pas à l'entrée. Elle peut valoir `adParamOutput` (sortie), `adParamInputOutput` (entrée/sortie) ou `adParamReturnValue` (valeur de retour). Ces directions prennent tout leur sens avec les **procédures stockées** : après `Execute`, on relit la valeur d'un paramètre de sortie pour récupérer un résultat calculé par le serveur.

En pratique, ce mécanisme concerne surtout **SQL Server** ; les requêtes ad hoc du moteur ACE n'exploitent pas de paramètres de sortie. L'appel de procédures stockées paramétrées est donc développé dans le contexte des bases externes (section 10.9).

---

## Command, requêtes sauvegardées et sécurité

Deux renvois utiles pour situer l'objet `Command` dans l'ensemble de la formation. D'une part, Access permet d'enregistrer des **requêtes paramétrées** dans la base elle-même (clause `PARAMETERS`) ; on peut les invoquer par code, sujet traité côté DAO au chapitre 12.3. D'autre part, l'objet `Command` est l'**arme anti-injection** d'ADO : c'est l'application concrète des principes de SQL paramétré exposés au chapitre 11.5, qui devrait être la règle dès qu'une valeur d'origine externe entre dans une requête.

---

## Quand utiliser Command plutôt que `Connection.Execute` ?

Tous les chemins d'exécution ne se valent pas selon le contexte. Pour une requête action **sans valeur variable**, `Connection.Execute` avec une chaîne reste le plus simple et suffit amplement. Pour une simple **lecture sans paramètre**, `Recordset.Open` avec une chaîne fait l'affaire. L'objet `Command` s'impose dès qu'apparaît l'un de ces besoins : **paramétrer** une requête (sécurité, formatage), la **réutiliser** efficacement, ou appeler une **procédure stockée**. En résumé, `Command` n'est pas un passage obligé pour tout, mais il devient incontournable dès que des valeurs variables ou la réutilisation entrent en jeu.

---

## En résumé

L'objet `Command` sépare la structure d'une requête de ses valeurs, apportant sécurité (contre l'injection SQL), correction (valeurs typées sans formatage) et réutilisation. Avec le moteur ACE, les paramètres s'écrivent avec le marqueur **positionnel `?`** — l'ordre prime, le nom est informatif, et les paramètres nommés `@nom` n'y ont pas cours. On définit chaque paramètre avec `CreateParameter` puis `Append`, en veillant à la **taille obligatoire** des types texte. La commande s'exécute via `Execute`, qui renvoie un recordset pour un `SELECT` ou le nombre de lignes affectées pour une requête action. Préparée (`Prepared = True`) et réexécutée en ne changeant que les valeurs, elle gagne en performance. Enfin, `Command` n'est requis que lorsque le paramétrage, la réutilisation ou les procédures stockées sont en jeu — pour le reste, `Connection.Execute` ou `Recordset.Open` suffisent.

Les sections précédentes supposaient toutes une connexion active. ADO offre pourtant une capacité que DAO n'a pas : travailler sur des données **après avoir fermé la connexion**. C'est la promesse des recordsets déconnectés, que nous explorons maintenant.


⏭️ [10.7. Recordsets déconnectés et persistance XML](/10-ado-access/07-recordsets-deconnectes.md)
