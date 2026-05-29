🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 18.4. Mise en cache des données et variables de module

L'accès le plus rapide à une donnée est celui que l'on n'effectue pas. Dans une application Access, chaque lecture peut solliciter le disque ou le réseau ; relire en permanence des informations qui ne changent pas est un gaspillage. La mise en cache consiste à lire une donnée stable une seule fois, à la conserver en mémoire, puis à la réutiliser — un levier de performance considérable, en particulier pour les données de référence consultées de façon répétée.

Mais ce gain a une contrepartie qu'il faut affronter de front : une donnée mise en cache peut devenir **périmée**. Toute la difficulté du cache tient dans l'arbitrage entre vitesse et fraîcheur, et dans la gestion de son invalidation.

## Le principe : lire une fois, réutiliser

Sont de bons candidats à la mise en cache les données qui changent rarement pendant une session : tables de référence (pays, codes de statut, taux de TVA), paramètres de configuration de l'application, rôles de l'utilisateur courant, ou encore le résultat d'un calcul coûteux que l'on rappellerait à l'identique.

À l'inverse, on ne met pas en cache — ou alors avec de grandes précautions — les données transactionnelles changeantes, ni celles que d'autres utilisateurs modifient en parallèle : un cache afficherait alors des valeurs dépassées. La règle directrice est d'aligner la durée de vie du cache sur la volatilité de la donnée : ce qui doit toujours être à jour ne se met pas en cache.

## Variables de module : une mémoire qui persiste

Dans un module standard, une variable déclarée au niveau module (`Private` ou `Public` en tête de module) conserve sa valeur d'un appel de procédure à l'autre, pour toute la durée de la session — contrairement à une variable locale, réinitialisée à chaque appel. C'est le support naturel d'un cache.

On l'alimente typiquement à la **première utilisation** (initialisation paresseuse) : la valeur est lue une fois, puis servie directement les fois suivantes.

```vba
Private mTauxTVA As Double
Private mTauxTVACharge As Boolean

Public Function TauxTVA() As Double
    If Not mTauxTVACharge Then
        mTauxTVA = Nz(DLookup("Valeur", "Parametres", "Cle = 'TauxTVA'"), 0.2)
        mTauxTVACharge = True
    End If
    TauxTVA = mTauxTVA
End Function
```

Le drapeau `mTauxTVACharge` n'est pas accessoire : un `Double` ne permet pas de distinguer « pas encore chargé » d'une valeur légitimement nulle. Sans indicateur séparé (ou valeur sentinelle), on relirait la source à chaque appel si la donnée valait 0.

> ℹ️ La portée et la durée de vie des variables sont détaillées au [chapitre 3.4](../03-rappels-fondamentaux/04-portee-variables.md).

## Variables `Static` : un cache local à une fonction

Lorsque le cache ne concerne qu'une seule fonction, le mot-clé `Static` offre une alternative plus compacte : une variable `Static` est locale à la procédure mais conserve sa valeur entre les appels.

```vba
Public Function TauxTVA() As Double
    Static valeur As Double
    Static charge As Boolean
    If Not charge Then
        valeur = Nz(DLookup("Valeur", "Parametres", "Cle = 'TauxTVA'"), 0.2)
        charge = True
    End If
    TauxTVA = valeur
End Function
```

L'effet est le même que précédemment, mais la valeur cachée reste encapsulée dans la fonction.

## Un cache de correspondances avec `Dictionary`

Le schéma le plus rentable concerne les recherches répétées. Appeler `DLookup` dans une boucle revient à lancer une requête à chaque tour — c'est un anti-pattern de performance majeur.

```vba
' À éviter : un DLookup par itération = une requête à chaque tour
Do Until rs.EOF
    nom = DLookup("Nom", "Clients", "ClientID = " & rs!ClientID)
    '  ...
    rs.MoveNext
Loop
```

La parade consiste à précharger la table de référence dans un `Scripting.Dictionary`, puis à effectuer des recherches en mémoire, quasi instantanées : une seule requête remplace les N appels.

```vba
Private mClients As Object        ' Scripting.Dictionary (liaison tardive)

Private Sub ChargerCacheClients()
    Dim rs As DAO.Recordset
    Set mClients = CreateObject("Scripting.Dictionary")
    Set rs = CurrentDb.OpenRecordset( _
        "SELECT ClientID, Nom FROM Clients", dbOpenForwardOnly)
    Do Until rs.EOF
        mClients(rs!ClientID.Value) = rs!Nom.Value
        rs.MoveNext
    Loop
    rs.Close
End Sub

Public Function NomClient(ByVal id As Long) As String
    If mClients Is Nothing Then ChargerCacheClients
    If mClients.Exists(id) Then NomClient = mClients(id)
End Function
```

On teste systématiquement `.Exists` avant de lire une clé : accéder directement à une clé absente d'un `Dictionary` l'ajoute silencieusement avec une valeur vide, ce qui finirait par polluer le cache.

> ℹ️ Le `Scripting.Dictionary` et les collections sont traités au [chapitre 3.7](../03-rappels-fondamentaux/07-collections-dictionary.md) ; les fonctions de domaine comme `DLookup` au [chapitre 11.10](../11-sql-access-vba/10-fonctions-domaine.md).

### Mais d'abord : une jointure ne suffirait-elle pas ?

Avant de mettre en cache, il faut se demander si une **jointure** ne réglerait pas le problème à la source. Si l'on parcourt une table et qu'on a besoin, pour chaque ligne, d'une information liée, la bonne réponse est souvent d'écrire une requête qui joint les deux tables — et non de boucler en interrogeant une table de référence. Le cache mémoire est pertinent pour des données transversales, réutilisées dans de nombreux endroits ou opérations ; une lecture isolée qui a besoin de données associées relève du raisonnement ensembliste vu aux [sections 18.2](02-optimisation-recordsets.md) et [18.3](03-execute-vs-runsql.md).

## Le vrai défi : l'invalidation du cache

Mettre en mémoire est facile ; savoir quand la mémoire ment l'est moins. Plusieurs stratégies permettent de gérer la péremption, selon le contexte.

La plus simple consiste à ne mettre en cache que des données **stables le temps de la session** — une configuration qui ne change pas tant que l'application tourne — et à reconstruire le cache au prochain démarrage.

Lorsqu'on sait qu'une donnée vient d'être modifiée, on prévoit une procédure de **rechargement explicite**, appelée à ce moment-là ou sur action de l'utilisateur.

On peut aussi recourir à une **expiration par le temps** : on mémorise l'instant du chargement et l'on recharge si le cache dépasse un certain âge.

```vba
Private mChargeLe As Date
Private Const CACHE_VALIDITE_MIN As Long = 10

Public Function NomClient(ByVal id As Long) As String
    If mClients Is Nothing _
       Or DateDiff("n", mChargeLe, Now) > CACHE_VALIDITE_MIN Then
        ChargerCacheClients          ' qui fait : mChargeLe = Now
    End If
    If mClients.Exists(id) Then NomClient = mClients(id)
End Function
```

En environnement multi-utilisateur enfin, on peut maintenir dans une petite table de contrôle un marqueur de version (ou un horodatage de dernière modification) que l'on consulte à moindre coût avant d'utiliser le cache : si le marqueur a changé, on recharge. Cette approche concilie fraîcheur et économie, en n'effectuant qu'une lecture légère plutôt qu'un rechargement complet.

> ℹ️ La gestion des accès concurrents est traitée au [chapitre 15](../15-multi-utilisateurs/README.md).

## Concevoir un cache qui se reconstruit seul

Un point souvent ignoré : les variables de module et `Static` perdent leur contenu lorsque l'état d'exécution du projet VBA est réinitialisé — par une erreur non gérée qui interrompt le code, par l'instruction `End`, ou lors d'une modification du code dans l'éditeur. Un cache ne doit donc jamais *supposer* qu'il est encore peuplé.

C'est précisément ce que garantit le schéma d'initialisation paresseuse : l'accesseur vérifie toujours l'état du cache (`Is Nothing`, drapeau de chargement) et le reconstruit si besoin. Le cache devient ainsi **auto-réparant** : si la mémoire a été perdue, le prochain accès la rétablit de façon transparente. On évite par ailleurs l'instruction `End`, qui efface brutalement tout l'état de l'application.

> ℹ️ Une bonne gestion des erreurs (sans `End`) est l'objet du [chapitre 13](../13-gestion-erreurs/README.md).

## Bien dimensionner le cache

La mise en cache vise des données de **petite à moyenne** taille. Charger une table d'un million de lignes dans un `Dictionary` gaspille la mémoire, et le chargement lui-même devient coûteux — l'inverse de l'objectif. On réserve donc le cache aux référentiels et paramètres compacts, et l'on s'en remet aux requêtes ciblées pour les gros volumes.

## `TempVars` et autres mémoires de session

Pour de petites valeurs globales (utilisateur courant, paramètre simple), Access offre les **TempVars** : des variables de session accessibles partout — y compris dans les requêtes et les macros — et plus robustes que les variables de module face aux réinitialisations, car gérées par Access et non par le projet VBA. Elles ne stockent toutefois que des valeurs scalaires, pas des objets ni des Recordsets.

Pour conserver tout un jeu d'enregistrements en mémoire, un **Recordset ADO déconnecté** peut aussi servir de cache structuré.

> ℹ️ Voir les TempVars au [chapitre 15.7](../15-multi-utilisateurs/07-tempvars.md) et les Recordsets déconnectés au [chapitre 10.7](../10-ado-access/07-recordsets-deconnectes.md).

## Centraliser dans un module dédié

Pour rester maintenable, on regroupe les structures de cache et leurs accesseurs dans un module standard dédié (par exemple `modCache`), assorti d'une routine `ReinitialiserCaches` qui remet tous les caches à `Nothing` (et les drapeaux à `False`). Le rechargement intervient alors automatiquement au prochain accès, et l'on dispose d'un point unique pour vider l'ensemble lorsque c'est nécessaire.

## Points clés à retenir

- Lire une fois et réutiliser : le cache mémoire évite les lectures répétées sur le disque ou le réseau.
- On met en cache des données stables (référentiels, paramètres) ; jamais des données qui doivent toujours être à jour.
- Les variables de module conservent leur valeur durant la session ; `Static` offre un cache local à une fonction. Un drapeau de chargement distingue « non chargé » d'une valeur nulle.
- Remplacer un `DLookup` répété en boucle par un `Dictionary` préchargé est l'un des gains les plus nets — après avoir vérifié qu'une jointure ne suffirait pas.
- L'invalidation est le vrai défi : limiter au périmètre de la session, recharger explicitement, expirer par le temps, ou suivre un marqueur de version en multi-utilisateur.
- Concevoir le cache comme auto-réparant (il se vérifie et se reconstruit), car l'état du projet peut être réinitialisé ; éviter `End`.
- Réserver le cache aux petits et moyens volumes ; envisager les `TempVars` pour les valeurs globales simples.

---


⏭️ [18.5. Optimisation des formulaires (RecordSource, filtres, chargement différé)](/18-optimisation-performance/05-optimisation-formulaires.md)
