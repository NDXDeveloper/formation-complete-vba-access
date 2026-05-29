🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 5.4. Exécution de requêtes (RunSQL, OpenQuery)

La section 5.2 a montré comment **ouvrir** une requête ; cette section se concentre sur la façon de l'**exécuter** — en particulier les requêtes action. `DoCmd` propose pour cela `RunSQL` et `OpenQuery`. Mais un fil rouge traverse cette section : pour exécuter une requête action **par code**, une autre voie — **`db.Execute`** (DAO) — est généralement **préférable**. Voyons les deux, et pourquoi.

## DoCmd.RunSQL — exécuter du SQL action

```
DoCmd.RunSQL SQLStatement, [UseTransaction]
```

`RunSQL` exécute une instruction SQL **action** fournie directement sous forme de chaîne — sans qu'il soit nécessaire d'avoir une requête sauvegardée :

```vba
DoCmd.RunSQL "UPDATE tblTarifs SET PrixHT = PrixHT * 1.02"
DoCmd.RunSQL "DELETE FROM tblTemp WHERE DateImport < #01/01/2025#"
```

Deux points à retenir :

- `RunSQL` ne traite que les requêtes **action** (`INSERT`, `UPDATE`, `DELETE`, création de table, DDL). On **ne peut pas** « exécuter » une requête `SELECT` ainsi pour en récupérer les résultats.
- Par défaut, `RunSQL` affiche les **messages de confirmation** (« Vous allez mettre à jour N ligne(s)… »), que l'on désactive avec `DoCmd.SetWarnings` (section 5.7) — non sans risque, comme on le verra.

À noter aussi que la chaîne SQL de `RunSQL` est soumise à une **limite de longueur** ; pour de très longues instructions, `db.Execute` est mieux adapté.

## DoCmd.OpenQuery — exécuter une requête sauvegardée

Comme vu en section 5.2, `OpenQuery` appliqué à une requête **action sauvegardée** l'**exécute** (en affichant, là encore, les messages de confirmation par défaut) :

```vba
DoCmd.OpenQuery "qryArchiver"      ' exécute la requête action sauvegardée
```

Rappel : pour une requête **SELECT**, `OpenQuery` n'« exécute » pas à proprement parler — il ouvre la **feuille de données** des résultats.

## L'alternative recommandée : db.Execute

Pour exécuter une requête action **par code**, la méthode **`Execute`** de l'objet `Database` (DAO, chapitre 11) est presque toujours le meilleur choix. Elle présente trois avantages décisifs sur `RunSQL`/`OpenQuery` :

- elle n'affiche **aucun message de confirmation** ;
- avec l'option **`dbFailOnError`**, elle **signale les erreurs** en déclenchant une erreur interceptable (violation de contrainte, par exemple) ;
- elle renseigne **`RecordsAffected`**, le nombre exact d'enregistrements touchés.

```vba
Dim db As DAO.Database
Set db = CurrentDb
db.Execute "UPDATE tblTarifs SET PrixHT = PrixHT * 1.02", dbFailOnError
Debug.Print db.RecordsAffected & " ligne(s) mise(s) à jour"
```

On retrouve ici l'idiome de la section 4.6 : `CurrentDb` est **capturé dans une variable** afin que `RecordsAffected` porte bien sur le **même** objet que celui qui a exécuté la requête.

## Le danger de SetWarnings False

Si l'on tient à utiliser `RunSQL` ou `OpenQuery`, on désactive souvent les messages avec `DoCmd.SetWarnings False`. **C'est un schéma risqué.** Si une **erreur survient** entre la désactivation et la réactivation, le code peut sortir **avant** d'avoir rétabli les messages — qui restent alors **désactivés pour toute la session Access**, masquant ensuite d'autres avertissements et confirmations.

```vba
' ❌ Schéma risqué : en cas d'erreur, SetWarnings True n'est jamais atteint
DoCmd.SetWarnings False
DoCmd.RunSQL "UPDATE tblTarifs SET PrixHT = PrixHT * 1.02"   ' si une erreur survient ici…
DoCmd.SetWarnings True                                        ' …cette ligne est ignorée

' ✅ db.Execute n'affiche aucun message : aucun SetWarnings à gérer
CurrentDb.Execute "UPDATE tblTarifs SET PrixHT = PrixHT * 1.02", dbFailOnError
```

C'est un argument supplémentaire — et de poids — en faveur de `db.Execute` : en l'absence de messages à supprimer, **le piège n'existe tout simplement pas**. Si l'on doit malgré tout employer `SetWarnings False`, il faut impérativement rétablir les messages dans un gestionnaire d'erreurs (chapitre 13), pour garantir leur réactivation quoi qu'il arrive.

## RunSQL/OpenQuery vs db.Execute — comparaison

| Critère | `DoCmd.RunSQL` / `OpenQuery` | `db.Execute` (DAO) |
|---|---|---|
| Type de requête | action | action |
| Messages de confirmation | oui (par défaut) | non |
| `SetWarnings` nécessaire | souvent | non |
| Nombre de lignes affectées | peu fiable | `db.RecordsAffected` (fiable) |
| Signalement des erreurs | limité | `dbFailOnError` → erreur interceptable |
| Recommandation | usage rapide / contexte UI | **à privilégier par code** |

La comparaison complète entre `Execute` et `RunSQL`, y compris les aspects de performance, est développée en section 18.3. L'exécution de SQL et les requêtes action sont, de manière générale, traitées au chapitre 11.

## À retenir

- **`DoCmd.RunSQL`** exécute une requête **action** fournie en chaîne SQL ; **`DoCmd.OpenQuery`** exécute une requête action **sauvegardée**. Aucune des deux ne « récupère » les résultats d'un `SELECT`.
- Toutes deux affichent par défaut des **messages de confirmation**, désactivables avec `SetWarnings` — au prix d'un risque.
- **`db.Execute "…", dbFailOnError`** est l'approche **recommandée par code** : pas de message, **erreurs interceptables** et **`RecordsAffected`** fiable (en capturant `CurrentDb` dans une variable, section 4.6).
- **Piège de `SetWarnings False`** : en cas d'erreur avant le `SetWarnings True`, les messages restent **désactivés pour toute la session** — raison forte de préférer `db.Execute`.
- Comparaison détaillée en section 18.3 ; exécution SQL au chapitre 11.

---


⏭️ [5.5. Import et export de données (TransferDatabase, TransferSpreadsheet, TransferText)](/05-objet-docmd/05-import-export-donnees.md)
