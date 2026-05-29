🔝 Retour au [Sommaire](/SOMMAIRE.md)

# Chapitre 18 — Optimisation et performance

Access traîne une réputation tenace : « ça rame ». Pourtant, dans l'immense majorité des cas, la lenteur d'une application Access ne vient pas du moteur **ACE** (Access Database Engine) lui-même, mais de la manière dont l'application a été conçue et codée. Une base correctement structurée, dont le code VBA exploite intelligemment les Recordsets, les requêtes ensemblistes et les index, peut servir confortablement plusieurs dizaines d'utilisateurs sur un réseau local. À l'inverse, quelques anti-patterns suffisent à transformer une opération d'une seconde en plusieurs minutes.

Ce chapitre adopte une approche résolument pragmatique de l'optimisation : on ne devine pas, on mesure ; on ne réécrit pas tout, on cible les véritables goulots d'étranglement. L'objectif n'est pas d'accumuler des « trucs » censés accélérer le code, mais de comprendre *où* le temps est réellement consommé dans une application Access, et *pourquoi*.

## Pourquoi l'optimisation mérite un chapitre à part

Access n'est pas une base de données client/serveur classique. Avec une base au format `.accdb`, le moteur ACE s'exécute sur le poste de chaque utilisateur, et non sur un serveur central. Quand l'application est découpée en front-end (interface) et back-end (données partagées sur le réseau), c'est le poste client qui exécute le traitement des requêtes : il rapatrie les données nécessaires — parfois bien plus que nécessaire — à travers le réseau.

Cette architecture change radicalement le « modèle de coût » d'une opération. Sur un véritable serveur SQL, le coût dominant est souvent le calcul côté serveur. Sur Access partagé, le coût dominant est presque toujours le **transfert de données et le nombre d'allers-retours réseau**. Optimiser une application Access consiste donc, avant tout, à réduire le volume de données déplacées et le nombre d'opérations atomiques, plutôt qu'à micro-optimiser des boucles VBA.

Comprendre cette particularité, c'est déjà éviter la plupart des erreurs de performance : beaucoup de techniques valables ailleurs n'ont ici qu'un effet négligeable, tandis que des décisions de conception apparemment anodines (un formulaire lié à une table entière, une boucle qui exécute une requête à chaque tour) ont un impact démesuré.

## La règle d'or : mesurer avant d'optimiser

L'optimisation prématurée est une perte de temps doublée d'un risque. Réécrire un fragment de code « pour qu'il aille plus vite » sans savoir s'il représente une part significative du temps d'exécution, c'est complexifier l'application sans bénéfice mesurable — et souvent y introduire des bugs.

La démarche saine est toujours la même : mesurer l'état actuel, identifier le ou les points qui concentrent l'essentiel du temps, optimiser ceux-là en priorité, puis mesurer de nouveau pour confirmer le gain. En pratique, une très petite fraction du code est responsable de la majeure partie des lenteurs perçues. C'est elle qu'il faut trouver, et c'est précisément l'objet de la première section de ce chapitre.

> ℹ️ Le profilage rejoint naturellement les techniques de débogage et d'instrumentation présentées au [chapitre 19 — Débogage et tests](../19-debogage-tests/README.md). Mesurer, c'est aussi savoir observer son code.

## Les véritables sources de lenteur dans Access

Avant d'entrer dans le détail, il est utile d'avoir en tête la carte des principaux goulots d'étranglement. Ils reviennent dans presque toutes les applications Access lentes :

Le **trafic réseau** est le premier suspect dès qu'une base est partagée. Réouvrir la connexion au back-end à chaque requête, rapatrier des tables entières ou maintenir des liaisons inefficaces multiplie les allers-retours. Maintenir une connexion persistante et limiter le volume transféré change souvent l'expérience du tout au tout.

L'**usage des Recordsets** vient ensuite. Choisir le mauvais type (Dynaset là où un Snapshot ou un Forward-only suffirait), parcourir des enregistrements un à un pour faire ce qu'une seule requête SQL réaliserait, ou ouvrir des jeux d'enregistrements bien plus larges que nécessaire : autant de causes fréquentes de lenteur côté code.

Le **mode d'exécution du SQL** a aussi son importance. Les traitements ligne par ligne (un `RunSQL` ou un `Update` de Recordset dans une boucle) sont presque toujours dramatiquement plus lents que la requête ensembliste équivalente, qui laisse le moteur faire le travail en une seule passe.

Les **formulaires** sont un point sensible de l'interface. Un formulaire lié à une table volumineuse sans filtre charge tout au démarrage ; multiplier les contrôles liés et les sous-formulaires alourdit encore l'ouverture. Un `RecordSource` ciblé et un chargement différé font une différence immédiatement perceptible par l'utilisateur.

L'**indexation** conditionne directement la rapidité des recherches, des jointures et des tris. Un index absent transforme une recherche instantanée en parcours complet de la table ; à l'inverse, un index superflu pénalise les écritures et gonfle le fichier. L'indexation est un compromis à piloter consciemment.

La **taille et l'état du fichier**, enfin, influent sur l'ensemble. Une base `.accdb` se « gonfle » avec le temps sous l'effet des suppressions, des modifications et des objets temporaires ; le compactage régulier reconstruit les index et récupère l'espace. Et au-delà de tout cela plane la limite structurelle des 2 Go par fichier, qu'il faut savoir anticiper.

## Ce que vous allez apprendre

Ce chapitre déroule ces thèmes dans un ordre logique, de la mesure jusqu'aux limites du moteur :

- **[18.1. Profilage et mesure des performances en VBA Access](01-profilage-mesure-performances.md)** — Instrumenter le code pour savoir où le temps est réellement consommé, avant toute optimisation.
- **[18.2. Optimisation des Recordsets — choix du type et des options](02-optimisation-recordsets.md)** — Sélectionner le bon type de Recordset (Table, Dynaset, Snapshot, Forward-only) et les options adaptées au besoin.
- **[18.3. Execute vs RunSQL — quand et pourquoi](03-execute-vs-runsql.md)** — Comprendre la différence de performance et de comportement entre `CurrentDb.Execute` et `DoCmd.RunSQL`.
- **[18.4. Mise en cache des données et variables de module](04-cache-variables-module.md)** — Éviter de relire en permanence des données stables en les conservant en mémoire.
- **[18.5. Optimisation des formulaires (RecordSource, filtres, chargement différé)](05-optimisation-formulaires.md)** — Réduire le coût d'ouverture et d'usage des formulaires liés.
- **[18.6. Indexation et impact sur les performances SQL](06-indexation-performances-sql.md)** — Tirer parti des index pour accélérer les requêtes tout en maîtrisant leur coût.
- **[18.7. Compactage automatique de la base par code](07-compactage-automatique.md)** — Lutter contre le gonflement du fichier et reconstruire les index par programme.
- **[18.8. Optimisation réseau pour les bases partagées (persistance de connexion)](08-optimisation-reseau.md)** — Réduire les allers-retours réseau et maintenir une connexion persistante au back-end.
- **[18.9. Limites de taille (2 Go) — stratégies de contournement](09-limites-taille-2go.md)** — Reconnaître la limite des 2 Go et adopter les stratégies pour la repousser ou la contourner.

> ℹ️ Lorsque l'optimisation interne ne suffit plus — volumétrie importante, forte concurrence — la migration des données vers un moteur client/serveur devient la solution de fond. Cette piste est traitée au [chapitre 23 — Migration et interopérabilité](../23-migration-interoperabilite/README.md).

## Prérequis

Ce chapitre suppose une bonne maîtrise des notions abordées précédemment, en particulier l'accès aux données via [DAO (chapitre 9)](../09-dao-data-access-objects/README.md) et [ADO (chapitre 10)](../10-ado-access/README.md), le [SQL dans Access VBA (chapitre 11)](../11-sql-access-vba/README.md), ainsi que le fonctionnement des [formulaires (chapitre 6)](../06-formulaires/README.md). La compréhension de l'architecture front-end / back-end, présentée au [chapitre 15 — Multi-utilisateurs et verrouillage](../15-multi-utilisateurs/README.md), est également utile pour les sections relatives au réseau.

## En résumé

Optimiser une application Access, ce n'est pas appliquer une liste de recettes, mais raisonner sur le coût réel des opérations dans une architecture où le réseau et le volume de données pèsent plus que le calcul brut. La discipline tient en quelques principes : mesurer avant d'agir, cibler les vrais goulots d'étranglement, préférer les traitements ensemblistes aux traitements ligne par ligne, et entretenir la base dans la durée. Les sections suivantes mettent en pratique chacun de ces principes.

---


⏭️ [18.1. Profilage et mesure des performances en VBA Access](/18-optimisation-performance/01-profilage-mesure-performances.md)
