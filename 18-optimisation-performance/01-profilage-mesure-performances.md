🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 18.1. Profilage et mesure des performances en VBA Access

On n'optimise bien que ce que l'on a mesuré. Avant de réécrire une seule ligne de code « pour aller plus vite », il faut savoir *où* le temps est réellement consommé et *combien* il l'est. Cette première section constitue donc la boîte à outils du chapitre : comment chronométrer une opération en VBA, comment passer d'une mesure globale à un profilage fin, et surtout comment mesurer juste — car une mesure mal conduite induit en erreur bien plus sûrement qu'une absence de mesure.

## Mesurer, ne pas supposer

L'intuition est un mauvais guide en matière de performance. Le code qui *semble* lent (une longue boucle visible à l'écran) n'est souvent pas le coupable ; le vrai goulot d'étranglement se cache fréquemment dans une opération discrète — une requête sans index, une connexion réseau rouverte à chaque tour. La seule façon de trancher est d'instrumenter le code et d'observer les chiffres.

Une mesure isolée a par ailleurs peu de valeur : les durées varient d'une exécution à l'autre selon la charge du poste, l'activité réseau ou l'état des caches. On exécute donc une opération plusieurs fois, et l'on raisonne sur des statistiques — typiquement la **moyenne** (vision d'ensemble) et le **minimum** (estimation du coût « pur », celui qui subit le moins d'interférences). Un écart important entre les deux signale une mesure instable, qu'il faut interpréter avec prudence.

## Choisir un chronomètre adapté

VBA et l'API Windows offrent plusieurs façons de mesurer le temps, dont la précision varie de plusieurs ordres de grandeur. Le choix dépend de la durée à mesurer.

| Outil | Résolution effective | Limite | Usage conseillé |
|---|---|---|---|
| `Timer` (VBA) | ~1/100 s au mieux | Remise à zéro à minuit | Mesures grossières, durées de plusieurs secondes |
| `GetTickCount` (API) | ~10 à 16 ms | `Long` négatif après ~24,9 j d'uptime | Mesures à la milliseconde, sans dépendance |
| `QueryPerformanceCounter` (API) | < 1 µs | Aucune en pratique | Référence pour les opérations brèves |

### La fonction `Timer` — simple mais limitée

`Timer` renvoie un `Single` correspondant au nombre de secondes écoulées depuis minuit. Elle ne nécessite aucune déclaration d'API, mais sa précision réelle est faible (de l'ordre du centième de seconde au mieux), ce qui la rend inadaptée aux durées courtes. Elle est aussi remise à zéro à minuit : une mesure qui chevauche cet instant produit un résultat aberrant.

```vba
Dim t As Single
t = Timer
'  ... opération à mesurer ...
Debug.Print "Durée : " & Format(Timer - t, "0.000") & " s"
```

Elle reste acceptable pour situer rapidement une opération longue (« est-ce 2 secondes ou 20 ? »), mais pas davantage.

### `GetTickCount` — la milliseconde via l'API

`GetTickCount` renvoie le nombre de millisecondes écoulées depuis le démarrage du système. Attention à un piège fréquent : si l'**unité** est la milliseconde, la **résolution** effective correspond à la granularité de l'horloge système, soit typiquement 10 à 16 ms. De plus, la valeur étant stockée dans un `Long` signé, elle devient négative après environ 24,9 jours de fonctionnement continu du poste. La variante `GetTickCount64` (renvoyant un `LongLong`) élimine ce problème.

```vba
#If VBA7 Then
    Private Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long
#Else
    Private Declare Function GetTickCount Lib "kernel32" () As Long
#End If
```

### `QueryPerformanceCounter` — la haute résolution

Pour mesurer des opérations brèves de façon fiable, c'est l'outil de référence. Couplé à `QueryPerformanceFrequency` (qui donne le nombre de « tics » par seconde), il offre une résolution inférieure à la microseconde.

L'idiome classique en VBA consiste à recevoir les compteurs dans des variables de type `Currency`. Ce type occupe 8 octets — exactement la taille attendue par l'API pour un entier 64 bits — ce qui évite de manipuler deux `Long` séparés. Le type `Currency` applique en interne un facteur d'échelle de 10000 ; mais comme le compteur **et** la fréquence subissent le même facteur, celui-ci s'annule dans la division `(fin − début) / fréquence` : le résultat en secondes est exact.

> ℹ️ La déclaration d'API et la compatibilité 32/64 bits (`PtrSafe`, `LongLong`) sont détaillées au [chapitre 22.1](../22-api-windows-integration-office/01-declaration-api-windows.md) ; l'annexe [H](../annexes/h-declarations-api-32-64-bits.md) fournit des déclarations prêtes à l'emploi.

## Un chronomètre réutilisable

Plutôt que de recopier des appels d'API partout, on encapsule la mesure dans un module de classe. Le code suivant constitue un chronomètre `clsStopWatch` fondé sur `QueryPerformanceCounter`, exposant le temps écoulé en secondes et en millisecondes.

```vba
' === Module de classe : clsStopWatch ===
Option Compare Database
Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" _
        (ByRef lpPerformanceCount As Currency) As Long
    Private Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" _
        (ByRef lpFrequency As Currency) As Long
#Else
    Private Declare Function QueryPerformanceCounter Lib "kernel32" _
        (ByRef lpPerformanceCount As Currency) As Long
    Private Declare Function QueryPerformanceFrequency Lib "kernel32" _
        (ByRef lpFrequency As Currency) As Long
#End If

Private m_freq    As Currency   ' Fréquence du compteur (tics/seconde)
Private m_debut   As Currency
Private m_fin     As Currency
Private m_enCours As Boolean

Private Sub Class_Initialize()
    ' La fréquence est constante : on la lit une seule fois
    QueryPerformanceFrequency m_freq
End Sub

Public Sub Demarrer()
    m_enCours = True
    QueryPerformanceCounter m_debut
End Sub

Public Sub Arreter()
    QueryPerformanceCounter m_fin
    m_enCours = False
End Sub

Public Property Get Secondes() As Double
    Dim courant As Currency
    If m_enCours Then
        QueryPerformanceCounter courant   ' lecture « à la volée » sans arrêter
    Else
        courant = m_fin
    End If
    If m_freq = 0 Then
        Secondes = 0          ' compteur indisponible : on évite la division par zéro
    Else
        Secondes = (courant - m_debut) / m_freq
    End If
End Property

Public Property Get Millisecondes() As Double
    Millisecondes = Secondes * 1000
End Property
```

Son utilisation est directe :

```vba
Dim chrono As clsStopWatch
Set chrono = New clsStopWatch

chrono.Demarrer
'  ... opération à mesurer ...
chrono.Arreter

Debug.Print "Durée : " & Format(chrono.Millisecondes, "0.000") & " ms"
```

Pour amortir le bruit de mesure, on répète l'opération et l'on agrège les résultats. Le canevas suivant exécute une opération `n` fois et restitue moyenne et minimum :

```vba
Public Sub MesurerNFois(ByVal n As Long)
    Dim chrono As clsStopWatch
    Dim i As Long, total As Double, mini As Double
    Set chrono = New clsStopWatch
    mini = 1E+30

    For i = 1 To n
        chrono.Demarrer
        '  ... opération à mesurer ...
        chrono.Arreter
        total = total + chrono.Millisecondes
        If chrono.Millisecondes < mini Then mini = chrono.Millisecondes
    Next i

    Debug.Print "Itérations : " & n
    Debug.Print "Moyenne    : " & Format(total / n, "0.000") & " ms"
    Debug.Print "Minimum    : " & Format(mini, "0.000") & " ms"
End Sub
```

## Mesurer correctement : les pièges à éviter

Disposer d'un chronomètre précis ne suffit pas ; encore faut-il que la mesure reflète la réalité. Plusieurs facteurs faussent couramment les résultats.

**Code interprété ou compilé.** Le projet exécuté dans l'éditeur, avec des points d'arrêt actifs ou en mode pas à pas, est nettement plus lent que le code compilé. Toute mesure faite en cours de débogage est dénuée de sens. On compile donc le projet (menu *Débogage > Compiler*) et l'on mesure une exécution normale, sans interruption.

**L'effet du cache.** Le premier accès à des données lit le disque ou le réseau et alimente les caches du système et du moteur ACE ; les exécutions suivantes sont alors bien plus rapides. Ignorer cet effet conduit à comparer une mesure « à froid » avec des mesures « à chaud ». La règle pratique est d'écarter la première itération (mise en chauffe) ou, mieux, de mesurer explicitement les deux cas : la première exécution représente l'expérience réelle d'un utilisateur ouvrant l'application, les suivantes le régime établi.

**`Debug.Print` à l'intérieur de la zone mesurée.** Écrire dans la fenêtre Exécution immédiate a un coût non négligeable. Placé dans une boucle serrée, il fausse complètement la mesure. On chronomètre *autour* de la boucle, jamais à l'intérieur, et l'on réserve les affichages au rapport final.

**Le rafraîchissement de l'écran.** Le repeint des formulaires et la mise à jour de l'interface pendant un traitement consomment du temps. Pour mesurer le seul « temps machine », on neutralise l'affichage avec `Application.Echo False` (et `DoCmd.Hourglass`). Mais ce temps d'affichage est aussi celui que ressent l'utilisateur : il faut donc distinguer le *temps de traitement* du *temps perçu*, et savoir lequel on cherche à optimiser.

**Conditions non représentatives.** Mesurer sur une copie locale mono-utilisateur contenant quelques centaines d'enregistrements ne dit rien des performances en production. Les durées doivent être relevées dans des conditions réalistes : volume de données réel, base partagée (front-end / back-end), réseau de production, plusieurs utilisateurs connectés. Une requête instantanée sur 200 lignes peut s'effondrer sur 200 000.

**La correction de noms automatique.** Les options de suivi et de correction automatique des noms (*Track / Perform Name AutoCorrect*, dans les options de la base de données active) ajoutent une surcharge à de nombreuses opérations et contribuent au gonflement du fichier. Activées par défaut, elles parasitent les mesures et pénalisent les performances : on les désactive dans une application destinée à la production.

## Du chronométrage global au profilage fin

Une fois l'opération coûteuse repérée, le réflexe est de la découper. On instrumente plusieurs points de passage à l'intérieur de la procédure pour identifier *quelle partie* concentre le temps, plutôt que de mesurer le bloc entier comme une boîte noire.

```vba
Dim chrono As clsStopWatch
Set chrono = New clsStopWatch

chrono.Demarrer
'  ... étape 1 : chargement ...
Debug.Print "Après chargement : " & Format(chrono.Millisecondes, "0.0") & " ms"

'  ... étape 2 : traitement ...
Debug.Print "Après traitement : " & Format(chrono.Millisecondes, "0.0") & " ms"

'  ... étape 3 : écriture ...
Debug.Print "Après écriture    : " & Format(chrono.Millisecondes, "0.0") & " ms"
```

Comme le chronomètre ci-dessus lit son temps écoulé « à la volée » sans s'arrêter, chaque ligne affiche le cumul depuis le départ. Les écarts entre points de mesure révèlent immédiatement l'étape dominante — celle sur laquelle concentrer l'effort, conformément au principe selon lequel une fraction réduite du code est responsable de l'essentiel des lenteurs.

## Journaliser les mesures

La fenêtre Exécution immédiate convient au poste de développement, mais pas au diagnostic sur le poste d'un utilisateur, où l'éditeur n'est pas surveillé. Dans ce cas, on consigne les durées dans une table dédiée (par exemple `tJournalPerf`, avec horodatage, nom de l'opération et durée) ou dans un fichier texte. On dispose ainsi d'un historique exploitable après coup, et l'on peut comparer les performances entre postes ou détecter une dégradation dans le temps.

> ℹ️ Les techniques de journalisation applicative et de débogage sur poste utilisateur sont approfondies aux chapitres [19.5](../19-debogage-tests/05-journalisation-evenements.md) et [19.7](../19-debogage-tests/07-debogage-distance.md). L'usage de `Debug.Print` et `Debug.Assert` est détaillé en [19.2](../19-debogage-tests/02-debug-print-assert.md).

## Mesurer les performances des requêtes

Les requêtes méritent une attention particulière, car elles sont souvent au cœur des lenteurs. Le temps d'une requête action s'obtient en chronométrant son exécution, et le nombre de lignes touchées via `RecordsAffected` — une métrique utile pour vérifier qu'une requête fait bien ce qu'on attend.

```vba
Dim chrono As clsStopWatch
Dim db As DAO.Database
Set db = CurrentDb
Set chrono = New clsStopWatch

chrono.Demarrer
db.Execute "UPDATE Commandes SET Statut = 'Traité' WHERE DateLivraison < Date()", dbFailOnError
chrono.Arreter

Debug.Print db.RecordsAffected & " lignes en " & _
            Format(chrono.Millisecondes, "0.0") & " ms"
```

> ℹ️ Le choix entre `Execute` et `RunSQL` a lui-même un impact mesurable sur la performance ; il fait l'objet de la [section 18.3](03-execute-vs-runsql.md).

Pour comprendre *pourquoi* une requête est lente, on peut demander au moteur ACE de révéler son **plan d'exécution** (index utilisés, stratégie de jointure) grâce à une fonctionnalité de diagnostic non documentée officiellement, *ShowPlan*. Elle s'active en ajoutant, sous la clé de registre du moteur (rubrique `Engines\Debug`), une valeur chaîne nommée `JETSHOWPLAN` à `ON`. Le moteur écrit alors dans un fichier texte `SHOWPLAN.OUT` la description du plan retenu pour chaque requête exécutée.

Cette technique est puissante mais à réserver à un poste de développement : il faut la désactiver après usage (valeur à `OFF`), car le fichier grossit vite et l'écriture ralentit l'exécution. Le chemin exact de la clé dépend de la version d'Access installée et doit être vérifié avant toute modification du registre.

> ℹ️ La lecture et l'écriture du registre Windows depuis VBA sont traitées au [chapitre 22.9](../22-api-windows-integration-office/09-registre-windows.md).

## La méthode de travail

L'instrumentation n'a de sens qu'inscrite dans une démarche itérative et disciplinée. On mesure l'état actuel pour établir une référence ; on identifie le point qui concentre l'essentiel du temps ; on n'optimise *qu'une seule chose* à la fois ; puis on mesure de nouveau pour confirmer — chiffres à l'appui — que la modification a bien apporté un gain, sans régression ailleurs. Modifier plusieurs éléments simultanément interdit d'attribuer le résultat, et fait courir le risque d'avoir complexifié le code sans bénéfice.

## Points clés à retenir

- La mesure précède toujours l'optimisation : on instrumente le code plutôt que de se fier à l'intuition.
- Le chronomètre se choisit selon la durée : `Timer` pour le grossier, `GetTickCount` pour la milliseconde, `QueryPerformanceCounter` pour les opérations brèves.
- Un module de classe encapsule proprement la mesure et se réutilise dans toute l'application.
- On répète les mesures et l'on raisonne sur la moyenne et le minimum, jamais sur une valeur unique.
- Les principaux facteurs de distorsion sont le mode débogage, l'effet de cache, `Debug.Print` dans la zone mesurée, le rafraîchissement de l'écran et des conditions non représentatives.
- Pour les requêtes, on chronomètre `Execute` et l'on peut analyser le plan d'exécution via *ShowPlan*.
- L'optimisation est un cycle : mesurer, cibler, modifier un seul point, re-mesurer.

---


⏭️ [18.2. Optimisation des Recordsets — choix du type et des options](/18-optimisation-performance/02-optimisation-recordsets.md)
