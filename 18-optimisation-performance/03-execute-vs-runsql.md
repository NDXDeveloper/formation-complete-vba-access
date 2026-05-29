🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 18.3. Execute vs RunSQL — quand et pourquoi

Pour exécuter une requête action en VBA — un `INSERT`, un `UPDATE`, un `DELETE` ou une instruction DDL — deux chemins coexistent : `DoCmd.RunSQL`, qui passe par l'objet `DoCmd` et la couche applicative d'Access, et `CurrentDb.Execute`, qui s'adresse directement au moteur via DAO. Le choix n'est pas qu'une question de style : il a des conséquences sur la performance, mais aussi — et surtout — sur la robustesse du traitement. Disons-le d'emblée : dans une application professionnelle, `Execute` est le choix par défaut.

Ces deux méthodes ne concernent que les requêtes **action** : ni l'une ni l'autre ne renvoie de jeu d'enregistrements. Pour un `SELECT`, on ouvre un Recordset (voir la [section 18.2](02-optimisation-recordsets.md)).

## `DoCmd.RunSQL` — le chemin par l'interface

`RunSQL` exécute une requête action en passant par la couche d'interface d'Access, celle-là même qui sert lorsqu'on lance une requête à la main. C'est ce qui explique ses deux principaux inconvénients.

### Les boîtes d'avertissement et le piège `SetWarnings`

Par défaut, `RunSQL` déclenche les messages de confirmation d'Access (« Vous êtes sur le point de modifier N lignes… »). Pour les supprimer, on encadre l'appel par `DoCmd.SetWarnings False` puis `True`.

```vba
DoCmd.SetWarnings False
DoCmd.RunSQL "UPDATE Articles SET PrixTTC = PrixHT * 1.2"
DoCmd.SetWarnings True        ' à ne surtout pas oublier
```

Ce schéma est fragile pour deux raisons. D'abord, `SetWarnings` agit sur un état **global** de la session : si une erreur survient entre les deux lignes et que le code s'interrompt, les avertissements restent désactivés pour toute la suite — y compris lors de manipulations manuelles ultérieures. Ensuite, `SetWarnings False` ne masque pas seulement la confirmation : il masque aussi les avertissements de **non-exécution** (lignes non ajoutées pour cause de violation de clé, règle de validation non respectée, conversion impossible). On peut alors croire qu'une opération a réussi alors qu'une partie des lignes a été silencieusement écartée.

### Aucun retour exploitable

`RunSQL` n'indique pas combien de lignes ont été affectées. Pour le savoir, il faut recourir à d'autres moyens, ce qui alourdit le code.

### Quelques limites supplémentaires

La chaîne SQL passée à `RunSQL` est par ailleurs limitée (de l'ordre de 32 000 caractères), là où `Execute` accepte des instructions plus longues. `RunSQL` accepte un second argument optionnel, `UseTransaction` (vrai par défaut), qui détermine si l'opération s'effectue dans une seule transaction.

> ℹ️ Le rôle de `RunSQL` et de `SetWarnings` est présenté aux chapitres [5.4](../05-objet-docmd/04-execution-requetes.md) et [5.7](../05-objet-docmd/07-autres-methodes-utiles.md).

## `CurrentDb.Execute` — le chemin direct

`Execute` s'adresse directement à l'objet `Database` de DAO, sans passer par l'interface. Il n'affiche donc **jamais** de boîte d'avertissement — `SetWarnings` devient inutile — et il est plus rapide. Surtout, il offre deux atouts décisifs en matière de fiabilité.

### `dbFailOnError` : la robustesse

L'option `dbFailOnError` demande au moteur d'annuler l'opération et de lever une erreur interceptable dès qu'une ligne ne peut être traitée. Sans elle, le moteur peut exécuter les lignes valides et ignorer silencieusement les autres ; avec elle, on est certain qu'un échec partiel ne passera pas inaperçu.

```vba
Dim db As DAO.Database
Set db = CurrentDb
On Error GoTo Gestion

db.Execute "UPDATE Articles SET PrixTTC = PrixHT * 1.2", dbFailOnError
Debug.Print db.RecordsAffected & " lignes mises à jour"
Exit Sub

Gestion:
    MsgBox "Échec de la mise à jour : " & Err.Description
```

C'est l'exact inverse du couple `RunSQL` + `SetWarnings False`, qui tend à *masquer* les échecs : ici, on les fait remonter.

> ℹ️ La structure de gestion d'erreurs `On Error GoTo` est détaillée au [chapitre 13.2](../13-gestion-erreurs/02-on-error-goto.md).

### `RecordsAffected` : connaître l'impact

Après un `Execute`, la propriété `RecordsAffected` de l'objet `Database` renvoie le nombre de lignes affectées — information précieuse pour journaliser, contrôler, ou confirmer à l'utilisateur. C'est quelque chose que `RunSQL` ne sait pas faire.

### Le piège de `CurrentDb`

Une subtilité essentielle accompagne `RecordsAffected` : `CurrentDb` renvoie une **nouvelle** instance de l'objet `Database` à chaque appel. Interroger `RecordsAffected` sur un `CurrentDb` distinct de celui qui a exécuté la requête renvoie donc… 0.

```vba
' Faux : deux instances différentes
CurrentDb.Execute "UPDATE ...", dbFailOnError
Debug.Print CurrentDb.RecordsAffected   ' renvoie 0 — autre objet Database

' Correct : une seule référence, conservée
Dim db As DAO.Database
Set db = CurrentDb
db.Execute "UPDATE ...", dbFailOnError
Debug.Print db.RecordsAffected
```

La règle est simple : on conserve une référence unique dans une variable `db` et l'on travaille toujours avec elle.

> ℹ️ La différence entre `CurrentDb()` et `DBEngine(0)(0)` est traitée au [chapitre 4.6](../04-modele-objet-access/06-currentdb-vs-dbengine.md).

## La différence qui surprend : la résolution des expressions

Il existe une différence de comportement souvent source de confusion. `RunSQL`, parce qu'il passe par Access, bénéficie du service d'expression de l'application : il sait résoudre une référence à un contrôle de formulaire écrite directement dans le SQL. `Execute`, lui, s'appuie sur le moteur ACE seul, qui ignore l'objet `Forms` : il interprète `Forms!frm!ctrl` comme un paramètre inconnu et échoue avec l'erreur 3061 (« Trop peu de paramètres. 1 attendu. »).

```vba
' Avec RunSQL : la référence au contrôle est résolue par Access
DoCmd.RunSQL "UPDATE Commandes SET Traite = True " & _
             "WHERE ClientID = Forms!frmClients!txtID"

' Avec Execute : ACE ne connaît pas Forms! -> erreur 3061
' Solution : concaténer la valeur dans la chaîne SQL
db.Execute "UPDATE Commandes SET Traite = True " & _
           "WHERE ClientID = " & Forms!frmClients!txtID, dbFailOnError
```

Cette différence ne doit pas faire pencher pour `RunSQL` : la concaténation (ou, mieux, une requête paramétrée) reste préférable. Elle impose toutefois de gérer correctement le formatage — délimiteurs de texte, format des dates et séparateur décimal — et de se prémunir contre l'injection SQL.

> ℹ️ Voir le SQL paramétré ([11.5](../11-sql-access-vba/05-sql-parametree.md)), la localisation du SQL dynamique ([11.7](../11-sql-access-vba/07-localisation-formatage-sql.md)) et les requêtes paramétrées par code ([12.3](../12-querydefs-tabledefs/03-requetes-parametrees.md)).

## Performance

En contournant la couche d'interface, `Execute` s'exécute plus rapidement que `RunSQL`, et il évite le va-et-vient sur `SetWarnings`. L'écart, modeste sur un appel isolé, devient significatif lorsque l'instruction est lancée de façon répétée — dans une boucle ou à chaque déclenchement d'un événement. À cet avantage de vitesse s'ajoutent la robustesse (`dbFailOnError`) et le retour d'information (`RecordsAffected`), qui font de `Execute` un choix supérieur sur tous les plans dans la grande majorité des cas.

## Et en ADO ?

En contexte ADO, l'équivalent est `Connection.Execute`, qui n'affiche pas davantage d'avertissements et renvoie le nombre de lignes affectées via un argument passé par référence.

```vba
Dim n As Long
CurrentProject.Connection.Execute "UPDATE Articles SET Actif = True", n, adCmdText
Debug.Print n & " lignes"
```

> ℹ️ Les opérations CRUD en ADO sont détaillées au [chapitre 10.5](../10-ado-access/05-crud-ado.md).

## Quand utiliser quoi

Par défaut, on utilise `Execute` : plus rapide, sans boîtes d'avertissement, avec contrôle des échecs via `dbFailOnError` et décompte via `RecordsAffected`. C'est l'option à privilégier dès qu'on écrit du nouveau code.

`RunSQL` ne se justifie guère que dans le cas où l'on tient à laisser Access résoudre des références de formulaire directement dans le SQL, sans refactoriser — et même alors, la concaténation ou les paramètres associés à `Execute` constituent une solution plus propre. En pratique, `RunSQL` est surtout l'héritage des macros converties en VBA ; on gagne à moderniser ces appels vers `Execute`.

> ℹ️ L'exécution du SQL par `RunSQL` et `Execute` est introduite au [chapitre 11.1](../11-sql-access-vba/01-execution-sql-runsql-execute.md) ; les requêtes DDL au [chapitre 11.4](../11-sql-access-vba/04-requetes-ddl.md).

## Points clés à retenir

- `RunSQL` et `Execute` n'exécutent que des requêtes action ; pour un `SELECT`, on ouvre un Recordset.
- `RunSQL` passe par l'interface : il déclenche des avertissements (d'où le recours fragile à `SetWarnings`), masque les échecs partiels et ne renvoie aucun décompte.
- `Execute` passe directement par DAO : aucun avertissement, exécution plus rapide, robustesse via `dbFailOnError` et nombre de lignes via `RecordsAffected`.
- Conserver une référence unique à `CurrentDb` dans une variable, sinon `RecordsAffected` renvoie 0.
- `RunSQL` résout les références `Forms!…` dans le SQL ; `Execute` non (erreur 3061). La parade est la concaténation ou les paramètres, pas le retour à `RunSQL`.
- Règle générale : `Execute` par défaut ; `RunSQL` réservé à l'héritage des macros, à moderniser.

---


⏭️ [18.4. Mise en cache des données et variables de module](/18-optimisation-performance/04-cache-variables-module.md)
