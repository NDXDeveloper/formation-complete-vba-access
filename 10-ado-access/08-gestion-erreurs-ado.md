🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 10.8. Gestion des erreurs ADO (objet Errors)

Toutes les sections précédentes ont supposé que les opérations réussissaient. En réalité, une connexion peut échouer, une requête être rejetée par le moteur, une contrainte d'intégrité être violée, un verrou entrer en conflit. Il faut donc savoir **intercepter et diagnostiquer** ces erreurs — et ADO offre pour cela un mécanisme qui lui est propre, distinct et plus riche que le simple objet `Err` de VBA : la **collection `Errors`**.

La particularité d'ADO tient à un fait souvent ignoré : **une seule opération ratée peut générer plusieurs erreurs à la fois**. Là où l'objet `Err` de VBA ne retient que la *dernière*, la collection `Errors` les conserve *toutes*, avec un niveau de détail propre au fournisseur. Cette section explique comment combiner les deux mécanismes pour un diagnostic complet. La gestion d'erreurs générale en VBA (structures `On Error`, journalisation) est traitée au chapitre 13, dont cette section constitue une application spécifique à ADO.

---

## Deux niveaux d'erreurs

Lorsqu'une opération ADO échoue, deux mécanismes entrent en jeu en parallèle.

D'un côté, l'**objet `Err` de VBA** : l'erreur est levée comme n'importe quelle erreur d'exécution et interceptée par un gestionnaire `On Error GoTo`. L'objet `Err` expose alors `Number`, `Description` et `Source`. C'est le mécanisme **universel** de VBA — il fonctionne toujours, mais il ne retient que **la dernière erreur**.

De l'autre, la **collection `Errors` d'ADO** : le fournisseur OLE DB peut signaler **plusieurs erreurs** pour un même incident — par exemple, un échec de connexion qui remonte simultanément une erreur réseau, une erreur du fournisseur et une erreur générique. Toutes sont rassemblées dans cette collection, avec des informations bien plus précises (code natif, état SQL) que ce que l'objet `Err` laisse transparaître.

> 💡 **L'idée clé** : `Err` répond à la question « *quelque chose a-t-il échoué, et quoi en dernier ?* », tandis que `Errors` répond à « *quelle est la liste complète et détaillée de ce qui a échoué, vue du fournisseur ?* ». Les deux sont complémentaires.

---

## L'objet Error et ses propriétés

Chaque élément de la collection est un objet **`Error`** (à ne pas confondre avec l'objet `Err` de VBA) doté de propriétés diagnostiques riches.

| Propriété | Contenu |
|---|---|
| `Number` | Numéro de l'erreur (code ADO ou code du fournisseur) |
| `Description` | Message décrivant l'erreur |
| `Source` | Objet ou fournisseur à l'origine de l'erreur |
| `SQLState` | Code d'état ANSI SQL (5 caractères), pour les sources SQL |
| `NativeError` | Code d'erreur **natif** du fournisseur (n° SQL Server, code moteur ACE…) |
| `HelpFile` / `HelpContext` | Fichier et contexte d'aide associés |

Les deux propriétés vraiment distinctives sont **`NativeError`** et **`SQLState`**. `NativeError` livre le code d'erreur *réel* du système sous-jacent — un numéro d'erreur SQL Server, par exemple — souvent bien plus parlant pour identifier la cause exacte que le code ADO générique. `SQLState`, lui, fournit le code d'état normalisé ANSI, utile face aux sources de données conformes SQL.

---

## La collection Errors : où vit-elle ?

Point essentiel : la collection `Errors` appartient à l'objet **`Connection`** — on y accède via `cn.Errors`. Elle est alimentée lorsqu'une erreur impliquant le **fournisseur** survient sur cette connexion. On la manipule comme une collection classique :

```vba
Debug.Print cn.Errors.Count          ' nombre d'erreurs
Debug.Print cn.Errors(0).Description ' accès par index
cn.Errors.Clear                      ' vide la collection
```

Une conséquence importante en découle : seules les erreurs **passant par le fournisseur** (échec de connexion, requête refusée, violation de contrainte…) garnissent cette collection. Une erreur purement *ADO* ou *VBA* — un argument invalide, une propriété en lecture seule mal employée — peut ne se manifester que via l'objet `Err`, **sans** alimenter `cn.Errors`. D'où la règle pratique : on s'appuie toujours sur `Err` (universel), et l'on consulte `cn.Errors` en complément dès qu'une connexion est en jeu.

---

## Le comportement de la collection

Deux comportements de la collection `Errors` doivent être connus pour l'utiliser correctement.

D'abord, elle est **vidée automatiquement** au début de l'opération ADO suivante susceptible de générer des erreurs : la nouvelle série remplace l'ancienne. Il faut donc la **lire immédiatement** après l'opération fautive, avant tout autre appel ADO — sous peine de la voir réinitialisée.

Ensuite, une opération peut **réussir tout en générant des avertissements** (et non des erreurs fatales). Dans ce cas, `cn.Errors` peut contenir des éléments alors qu'**aucune erreur VBA n'a été levée**. Autrement dit, `Errors.Count > 0` ne signifie pas systématiquement « échec » : il peut s'agir de simples messages d'information du fournisseur.

---

## Le patron de référence : `On Error GoTo` + énumération de `Errors`

En pratique, on combine les deux mécanismes : le gestionnaire `On Error GoTo` capte l'erreur, puis on **énumère `cn.Errors`** pour obtenir le détail complet du fournisseur. Le patron ci-dessous illustre cette approche.

```vba
Public Sub ExecuterAvecDiagnostic()
    Dim cn As ADODB.Connection
    Dim errADO As ADODB.Error

    On Error GoTo Gestion

    Set cn = New ADODB.Connection
    cn.Open "Provider=Microsoft.ACE.OLEDB.12.0;" & _
            "Data Source=" & CurrentProject.Path & "\Donnees.accdb;"

    ' ... opérations sur la connexion ...

Nettoyage:
    On Error Resume Next
    If Not cn Is Nothing Then
        If cn.State = adStateOpen Then cn.Close
        Set cn = Nothing
    End If
    Exit Sub

Gestion:
    ' 1) L'erreur VBA (toujours disponible)
    Debug.Print "Erreur VBA " & Err.Number & " : " & Err.Description

    ' 2) Le détail ADO, si une connexion existe
    If Not cn Is Nothing Then
        For Each errADO In cn.Errors
            Debug.Print "  ADO #" & errADO.Number & _
                        " | Natif=" & errADO.NativeError & _
                        " | SQLState=" & errADO.SQLState & _
                        " | " & errADO.Description & _
                        " | Source=" & errADO.Source
        Next errADO
    End If

    Resume Nettoyage
End Sub
```

Trois points méritent attention. On lit d'abord l'objet `Err`, **toujours** renseigné quelle que soit la nature de l'erreur. On énumère ensuite `cn.Errors` en se protégeant par `If Not cn Is Nothing`, car une erreur survenue *avant* l'existence de la connexion rendrait cet accès impossible. Enfin, l'énumération se fait **dans le gestionnaire, immédiatement**, avant tout nouvel appel ADO qui viderait la collection.

---

## Erreurs ou avertissements : l'événement `InfoMessage`

Comment capter les avertissements évoqués plus haut — ces messages que le fournisseur émet sans lever d'erreur fatale ? On peut bien sûr inspecter `cn.Errors.Count` après une opération, mais ADO offre un mécanisme dédié : l'événement **`InfoMessage`** de l'objet `Connection`. En déclarant la connexion `WithEvents` dans un module de classe, on intercepte ces messages au fil de l'eau, sans attendre une erreur. Ce mécanisme événementiel relève de la programmation orientée objet abordée au chapitre 16.4.

---

## Parallèle avec DAO

Le mécanisme n'est pas sans rappeler celui de DAO, qui maintient lui aussi sa propre collection d'erreurs **`DBEngine.Errors`**, distincte de l'objet `Err` de VBA. La logique est identique dans les deux technologies : le moteur (ACE pour DAO, le fournisseur pour ADO) recense ses erreurs dans une collection détaillée, tandis que VBA n'en expose que la dernière via `Err`. Ce parallèle, et la manière d'intercepter spécifiquement les erreurs de chaque technologie, sont traités au chapitre 13.5.

---

## En pratique : quand énumérer `Errors` ?

Inutile de déployer l'artillerie complète partout. Pour du **code courant**, la structure `On Error GoTo` associée à l'objet `Err` (chapitre 13.2) suffit amplement : on intercepte, on journalise le numéro et la description, on agit en conséquence.

L'énumération de `cn.Errors` prend en revanche tout son intérêt pour **diagnostiquer les échecs de connexion ou de requête**, en particulier face à des **serveurs externes** (section 10.9). C'est là que `NativeError` et `SQLState` font la différence : ils pointent la cause réelle — un numéro d'erreur SQL Server précis, un code du moteur ACE — là où le message VBA reste vague. Dans ces situations, on a tout intérêt à **journaliser la série complète** des erreurs (chapitre 13.6), et non la seule dernière.

À titre d'illustration, les incidents qui remontent typiquement par ce canal incluent le fournisseur non inscrit (le problème 32/64 bits de la section 10.2), un fichier introuvable ou verrouillé, un conflit de verrou en multi-utilisateur (chapitre 15), une incompatibilité de type sur un paramètre, ou une violation de contrainte lors d'une insertion. Les codes d'erreur courants d'Access, DAO et ADO sont par ailleurs catalogués dans l'annexe C.

---

## En résumé

ADO superpose deux mécanismes d'erreurs : l'objet **`Err` de VBA**, universel mais limité à la dernière erreur, et la collection **`Errors` de la connexion**, qui rassemble *toutes* les erreurs d'un incident avec un détail propre au fournisseur — notamment `NativeError` et `SQLState`. Cette collection appartient à l'objet `Connection`, n'est alimentée que par les erreurs passant par le fournisseur, et se vide à l'opération suivante : il faut donc la lire immédiatement. Le patron de référence combine `On Error GoTo` (pour capter) et l'énumération de `cn.Errors` (pour diagnostiquer en profondeur). Pour le code courant, `Err` suffit ; pour les échecs de connexion ou de requête, surtout côté serveur externe, l'énumération de `Errors` révèle la cause réelle.

Nous disposons maintenant de toutes les briques d'ADO : connexion, recordset, command paramétré, recordsets déconnectés et diagnostic des erreurs. Il est temps de mettre tout cela au service de ce qui constitue la vocation première d'ADO — se connecter à des bases de données *externes*. C'est l'objet de la dernière section du chapitre.


⏭️ [10.9. Connexion ADO à des bases externes (SQL Server, Oracle, PostgreSQL, MySQL)](/10-ado-access/09-connexion-bases-externes.md)
