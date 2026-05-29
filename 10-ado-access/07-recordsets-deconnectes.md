🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 10.7. Recordsets déconnectés et persistance XML

Voici la fonctionnalité qui distingue le plus nettement ADO de DAO, annoncée à plusieurs reprises dans ce chapitre : la capacité à **travailler sur des données après avoir fermé la connexion**. Un recordset DAO reste indéfectiblement attaché à sa base ; un recordset ADO, lui, peut être *déconnecté* — chargé en mémoire, puis détaché de sa source, libre d'être parcouru, trié, filtré, voire modifié sans qu'aucune connexion ne demeure ouverte.

Cette section explore deux fonctionnalités complémentaires : les **recordsets déconnectés** (manipulation de données en mémoire) et la **persistance** (sauvegarde et rechargement d'un recordset vers un fichier, notamment au format XML). Ensemble, elles ouvrent des usages impossibles en DAO : réduire la durée des verrous sur une base partagée, mettre des données en cache localement, ou transmettre un jeu d'enregistrements entre composants sans maintenir de connexion.

---

## Le concept de recordset déconnecté

Dans le fonctionnement normal, un recordset entretient un **lien vivant** avec sa source : chaque navigation, chaque modification dialogue avec le moteur. Un recordset déconnecté rompt ce lien. On charge les données, on **détache** le recordset de sa connexion (`ActiveConnection = Nothing`), puis on peut fermer la connexion — le recordset survit, autonome, en mémoire.

Plusieurs besoins motivent cette approche. **Réduire la durée des verrous** sur une base partagée : on charge, on se déconnecte aussitôt, on travaille tranquillement, et l'on ne se reconnecte qu'au moment de sauvegarder (un enjeu majeur en multi-utilisateur, chapitre 15). **Travailler hors connexion**, dans des scénarios où la base n'est pas toujours accessible. **Mettre des données de référence en cache** localement, pour éviter de réinterroger sans cesse le serveur. **Transmettre un recordset entre couches** d'une application, sans imposer à chacune de tenir une connexion (architecture en couches, chapitre 16.8).

> 💡 La déconnexion repose sur le **service de curseur côté client** d'ADO (le *Cursor Service for OLE DB*), une couche indépendante du fournisseur. Elle fonctionne donc avec ACE comme avec SQL Server ou tout autre fournisseur OLE DB.

---

## Créer un recordset déconnecté

La déconnexion impose deux prérequis stricts, à respecter **avant** l'ouverture.

Le premier est l'**emplacement du curseur côté client** : `CursorLocation = adUseClient`. Un curseur serveur ne peut jamais être déconnecté ; cette condition est non négociable. Le second, lorsqu'on entend modifier puis resynchroniser les données, est le **verrouillage par lot** : `adLockBatchOptimistic` (section 10.4). Une fois le recordset ouvert, la déconnexion proprement dite consiste à affecter `Nothing` à sa propriété `ActiveConnection`.

```vba
Dim cn As ADODB.Connection
Dim rs As ADODB.Recordset

Set cn = New ADODB.Connection
cn.Open "Provider=Microsoft.ACE.OLEDB.12.0;" & _
        "Data Source=" & CurrentProject.Path & "\Donnees.accdb;"

Set rs = New ADODB.Recordset
rs.CursorLocation = adUseClient                       ' IMPÉRATIF, avant Open
rs.Open "SELECT * FROM Clients", cn, adOpenStatic, adLockBatchOptimistic, adCmdText

Set rs.ActiveConnection = Nothing                     ' déconnexion du recordset

cn.Close: Set cn = Nothing                            ' la connexion peut désormais être fermée

' ... le recordset rs reste pleinement exploitable en mémoire ...
```

À partir de cet instant, `rs` vit sa propre vie : on le parcourt, on le lit, on le modifie, alors même qu'aucune connexion n'existe plus.

---

## Travailler en mémoire : trier et filtrer côté client

Un atout précieux des recordsets côté client (déconnectés ou non) est la possibilité de **trier et filtrer les données en mémoire**, sans réinterroger la base. La propriété **`Sort`** réordonne les enregistrements :

```vba
rs.Sort = "NomClient ASC"           ' tri en mémoire, sans nouvelle requête
rs.Sort = "Region ASC, Ville DESC"  ' tri multi-critères
rs.Sort = ""                        ' annule le tri
```

La propriété **`Filter`** restreint les enregistrements visibles, soit selon un critère textuel, soit à l'aide de constantes spéciales :

```vba
rs.Filter = "Region = 'Nord'"       ' filtre par critère
rs.Filter = adFilterNone            ' supprime le filtre (équivaut à "")
```

| Valeur de `Filter` | Effet |
|---|---|
| Chaîne de critère (`"Ville = 'Lyon'"`) | Filtre selon une condition |
| `adFilterNone` | Supprime tout filtre |
| `adFilterPendingRecords` | Enregistrements modifiés non encore appliqués |
| `adFilterConflictingRecords` | Enregistrements en échec après un `UpdateBatch` (conflit) |
| `adFilterAffectedRecords` | Enregistrements affectés par la dernière opération de lot |

Ces opérations, instantanées car effectuées en mémoire, évitent autant d'allers-retours avec le moteur — un gain net lorsque les mêmes données sont triées et filtrées de multiples façons.

---

## Modifier hors connexion et resynchroniser : `UpdateBatch`

C'est ici que le verrouillage `adLockBatchOptimistic` prend tout son sens. Sur un recordset déconnecté ouvert dans ce mode, les modifications — éditions, ajouts (`AddNew`), suppressions (`Delete`) — sont **mises en cache** dans le recordset, sans toucher à la source. Pour les appliquer, il faut **se reconnecter** puis appeler **`UpdateBatch`**, qui transmet toutes les modifications en attente en un seul lot.

```vba
' (rs a été déconnecté, modifié en mémoire...)
rs.Fields("Ville").Value = "Bordeaux"   ' modification mise en cache
rs.MoveNext
rs.Fields("Actif").Value = False        ' autre modification en cache

' Reconnexion pour appliquer le lot
Set cn = New ADODB.Connection
cn.Open "Provider=Microsoft.ACE.OLEDB.12.0;" & _
        "Data Source=" & CurrentProject.Path & "\Donnees.accdb;"
Set rs.ActiveConnection = cn

rs.UpdateBatch                          ' applique toutes les modifications en attente
```

La méthode symétrique **`CancelBatch`** abandonne les modifications en cache sans les appliquer. Pour qu'`UpdateBatch` réussisse, la requête d'origine doit être **modifiable** (table unique ou jointure modifiable dotée d'une clé unique) : un `SELECT` complexe non modifiable ne pourra pas être resynchronisé.

---

## Gérer les conflits à la resynchronisation

Entre le chargement et l'`UpdateBatch`, un autre utilisateur a pu modifier les mêmes enregistrements. ADO **détecte ces conflits** lors de l'application du lot et marque les enregistrements concernés. On les isole en filtrant sur `adFilterConflictingRecords`, et l'on examine l'état de chacun via la propriété **`Status`** (qui peut valoir `adRecOK`, `adRecModified`, `adRecConcurrencyViolation`, etc.) :

```vba
rs.Filter = adFilterConflictingRecords
Do While Not rs.EOF
    Debug.Print "Conflit sur l'enregistrement, statut : " & rs.Status
    rs.MoveNext
Loop
rs.Filter = adFilterNone
```

La résolution des conflits — choisir entre la version locale et la version distante, fusionner, abandonner — relève d'une stratégie applicative détaillée au chapitre 14.5. Retenez ici qu'ADO **signale** les conflits et **vous laisse décider**, plutôt que d'écraser aveuglément.

---

## Fabriquer un recordset de toutes pièces

Au-delà de la déconnexion d'un recordset issu d'une base, ADO permet de **construire un recordset entièrement en mémoire**, sans aucune source de données. On définit ses champs via la collection `Fields`, on l'ouvre, puis on le peuple :

```vba
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
rs.Fields.Append "ID", adInteger
rs.Fields.Append "Nom", adVarWChar, 50
rs.CursorLocation = adUseClient
rs.Open                                  ' aucun Source, aucune connexion

rs.AddNew Array("ID", "Nom"), Array(1, "Dupont")
rs.Update
rs.AddNew Array("ID", "Nom"), Array(2, "Martin")
rs.Update
```

Ce recordset « fabriqué » se comporte comme n'importe quel autre : on peut le trier, le filtrer, le parcourir. Il s'avère utile pour structurer des données calculées, alimenter un contrôle (liste, grille) sans table sous-jacente, ou servir de conteneur d'échange entre procédures.

---

## La persistance : sauvegarder et recharger

Un recordset peut être **sérialisé dans un fichier**, puis rechargé ultérieurement — y compris sans aucune connexion. C'est le rôle de la méthode **`Save`** et du rechargement par `Open` avec l'option `adCmdFile`.

### Sauvegarder avec `Save`

`Save` accepte un chemin et un format. Deux formats existent : **`adPersistADTG`** (format binaire propriétaire, compact et rapide) et **`adPersistXML`** (format XML, lisible et interopérable).

```vba
' Sauvegarde au format XML (lisible, échangeable)
rs.Save CurrentProject.Path & "\export.xml", adPersistXML

' Sauvegarde au format binaire ADTG (compact, rapide)
rs.Save CurrentProject.Path & "\cache.adtg", adPersistADTG
```

> ⚠️ **Piège du fichier existant** : `Save` **échoue** si le fichier cible existe déjà. Il faut donc le supprimer au préalable, par exemple : `If Dir(chemin) <> "" Then Kill chemin` avant l'appel à `Save`.

### Recharger avec `Open`

Le rechargement se fait sur un nouveau recordset, **sans connexion** (le deuxième argument est omis), en précisant `adCmdFile` :

```vba
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
rs2.Open CurrentProject.Path & "\export.xml", , adOpenStatic, adLockBatchOptimistic, adCmdFile
' rs2 est immédiatement exploitable, aucune base n'est requise
```

### XML ou ADTG : lequel choisir ?

Le choix dépend de l'usage.

| Format | Constante | Caractéristiques |
|---|---|---|
| Binaire ADTG | `adPersistADTG` | Compact, rapide à lire/écrire, **propriétaire** |
| XML | `adPersistXML` | Lisible, **interopérable**, mais plus volumineux |

Privilégiez **XML** lorsque le fichier doit être inspecté à l'œil, échangé avec un autre outil ou conservé dans un format ouvert et pérenne. Préférez **ADTG** pour un cache purement local, où compacité et rapidité priment sur la lisibilité. Le format XML produit un document décrivant à la fois la **structure** (schéma) et les **données**, ce qui le rend autonome et exploitable par des tiers.

À noter enfin que `Save` peut aussi écrire vers un objet **`Stream`** en mémoire plutôt que vers un fichier — pratique pour transmettre un recordset sur le réseau ou l'embarquer dans un autre flux, sans passage par le disque.

---

## Cas d'usage et limites

Les recordsets déconnectés et la persistance brillent dans plusieurs situations : **scénarios hors connexion** ou à connexion intermittente, **réduction de la durée des verrous** sur les bases partagées, **mise en cache** de données de référence rechargées sans solliciter le serveur, **transmission de données entre couches** applicatives, et alimentation de contrôles à partir de recordsets fabriqués.

Ces atouts s'accompagnent toutefois de limites à connaître. La **mémoire** : un recordset déconnecté conserve toutes ses données côté client — il est inadapté aux très gros volumes. La nature **statique** : c'est un instantané, qui ne reflète aucun changement survenu après le chargement. La **modifiabilité de la requête** : seul un `SELECT` modifiable (table unique ou jointure dotée d'une clé) pourra être resynchronisé par `UpdateBatch`. Et la **gestion des conflits**, qui devient votre responsabilité dès lors que plusieurs utilisateurs travaillent sur les mêmes données.

---

## En résumé

ADO permet de **déconnecter** un recordset de sa source pour le manipuler en mémoire, à condition de l'avoir ouvert côté client (`adUseClient`) et, pour l'édition différée, en verrouillage par lot (`adLockBatchOptimistic`) ; la déconnexion s'opère par `Set rs.ActiveConnection = Nothing`. Détaché, le recordset se trie et se filtre instantanément (`Sort`, `Filter`) sans réinterroger la base. Les modifications, mises en cache, s'appliquent par `UpdateBatch` après reconnexion, ADO signalant les éventuels conflits via `adFilterConflictingRecords` et la propriété `Status`. On peut aussi **fabriquer un recordset de toutes pièces** en mémoire, et **persister** n'importe quel recordset dans un fichier — XML pour l'interopérabilité, ADTG pour la compacité — rechargeable ensuite sans connexion. Autant de possibilités étrangères à DAO, à employer en gardant à l'esprit leurs limites de mémoire, de fraîcheur et de gestion des conflits.

Tout au long de ce chapitre, nous avons supposé que les opérations réussissaient. Mais une connexion peut échouer, une requête être rejetée, un conflit survenir. ADO dispose d'un mécanisme d'erreurs qui lui est propre, plus riche que le simple objet `Err` de VBA — c'est l'objet de la section suivante.


⏭️ [10.8. Gestion des erreurs ADO (objet Errors)](/10-ado-access/08-gestion-erreurs-ado.md)
