🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 15.10. Limites du moteur ACE en environnement concurrent

Tout au long de ce chapitre, une réalité est revenue comme un fil rouge : le moteur ACE est un **moteur fichier**, et non un système client-serveur. Cette nature fixe des limites bien réelles à ce qu'une application Access partagée peut accomplir. Cette dernière section les rassemble, sans alarmisme — ACE reste parfaitement adapté à de petits groupes de travail — mais avec lucidité, pour savoir jusqu'où aller et reconnaître le moment où une autre architecture s'impose.

## La nature fondamentale : moteur fichier, pas client-serveur

La limite première dont découlent toutes les autres : il n'existe **aucun processus serveur** qui arbitre les accès. Chaque poste ouvre directement le fichier partagé à travers le réseau, et **tout le traitement des requêtes se déroule côté client**. Lorsqu'un utilisateur exécute une requête, son poste rapatrie les données par le réseau pour les traiter localement — là où un véritable serveur de base de données traiterait la requête de son côté et ne renverrait que le résultat.

Cette différence d'architecture explique l'essentiel des limites qui suivent : dépendance au réseau, plafond de concurrence, et fragilité face aux interruptions.

## La limite du nombre d'utilisateurs concurrents

Le moteur supporte au maximum **255 utilisateurs simultanés** (rappel de la [section 15.8](08-gestion-sessions-utilisateurs.md), où l'on a vu que le fichier de verrouillage ne stocke pas plus de 255 entrées). Mais ce plafond théorique est trompeur : pour une application où les utilisateurs **ajoutent et modifient fréquemment** des données, il est recommandé de ne pas dépasser **25 à 50 utilisateurs**. Bien avant d'atteindre 255, la contention sur les verrous, le trafic réseau et la coordination via le fichier de verrouillage dégradent les performances de façon non linéaire. Pour une application essentiellement de consultation, on peut viser plus haut ; pour une application à écritures intensives, sensiblement moins.

## La limite de taille (2 Go)

Une base Access est limitée à **2 Go**. Dans une architecture front-end / back-end, c'est le **back-end** qui concentre toutes les données, et c'est donc lui qui approche cette limite — d'autant plus vite que les utilisateurs sont nombreux et actifs. À l'approche de 2 Go, les risques de corruption et de dégradation des performances augmentent nettement. Les stratégies de contournement (archivage, fractionnement du back-end, etc.) sont traitées à la [section 18.9](/18-optimisation-performance/09-limites-taille-2go.md).

## La dépendance et la fragilité réseau

Puisque les données transitent intégralement par le réseau, celui-ci devient le **point névralgique** de l'application. Sa latence et sa bande passante conditionnent directement les performances, et surtout, sa fiabilité conditionne l'**intégrité** des données.

C'est le point le plus critique en pratique : une **interruption réseau pendant une écriture** est la principale cause de corruption d'une base Access. Une connexion Wi-Fi instable, un VPN, un réseau local défaillant, une coupure de courant, voire un antivirus ou un pare-feu qui interfère avec le trafic, peuvent laisser le fichier dans un état incohérent. En conséquence, on recommande fortement une **connexion filaire** pour accéder au back-end, et l'on **proscrit le Wi-Fi et le VPN** pour cet usage. Le partage d'un back-end Access à travers un WAN ou Internet n'est pas viable.

## La limite de verrous par fichier (MaxLocksPerFile)

Le moteur impose un nombre maximal de verrous posés simultanément sur un fichier, défini par le réglage **`MaxLocksPerFile`**, dont la valeur par défaut est **9 500**. Lorsqu'une opération — typiquement une grosse transaction ou une requête action portant sur de nombreux enregistrements — dépasse ce nombre, le moteur lève l'**erreur 3052 « File sharing lock count exceeded »**.

Deux remarques importantes. D'abord, cette erreur n'est **pas strictement liée au multi-utilisateur** : une transaction volumineuse dans une seule session peut l'atteindre. Ensuite, la valeur peut être augmentée, soit dans le registre, soit temporairement pour la session courante via `DBEngine.SetOption dbMaxLocksPerFile, valeur`. Mais relever la limite traite le symptôme, pas la cause : il est souvent préférable de **fractionner** les grosses opérations en lots plus petits, validés séparément, plutôt que de gonfler indéfiniment le nombre de verrous.

## Un verrouillage coopératif, sans gestionnaire central

La concurrence d'ACE repose sur un verrouillage **coopératif** — verrous gérés via le fichier de verrouillage et plages d'octets sur le fichier partagé — et non sur un gestionnaire de verrous central comme en possède un serveur. Ce modèle fonctionne, mais il est moins robuste sous forte charge. Il s'accompagne d'autres contraintes déjà rencontrées : absence de **niveaux d'isolation** configurables (cf. [section 14.7](/14-transactions/07-niveaux-isolation.md)) et contention liée au **verrouillage de page** lorsque le verrouillage au niveau enregistrement n'est pas en vigueur (cf. [section 15.3](03-strategies-verrouillage.md)).

## Pas de client léger ni de web natif

Enfin, ACE suppose la présence du **client Access** (ou de son runtime) sur chaque poste et l'accès à un **partage de fichiers**. Ce n'est pas un moteur web ou cloud : il n'offre pas d'accès navigateur ni de modèle client léger. Pour un usage web ou mobile, on se tourne vers d'autres plateformes (cf. [sections 23.5](/23-migration-interoperabilite/05-migration-sharepoint-dataverse.md) et [23.6](/23-migration-interoperabilite/06-migration-power-apps.md)).

## Atténuer dans les limites d'ACE

Avant d'envisager une migration, plusieurs pratiques repoussent les limites et fiabilisent une application partagée :

- **réseau filaire**, sans Wi-Fi ni VPN pour le back-end ;
- **maintenir le back-end nettement sous 2 Go** : archiver les données anciennes, fractionner si nécessaire ;
- **verrouillage au niveau enregistrement, verrouillage optimiste et transactions courtes** (sections 15.3 et 14.x) pour réduire la contention ;
- **fractionner les grosses opérations** en lots pour rester sous `MaxLocksPerFile` ;
- **connexion persistante** au back-end et optimisation réseau ([section 18.8](/18-optimisation-performance/08-optimisation-reseau.md)) ;
- **compactage régulier** pour contenir le gonflement et préserver l'intégrité ([section 18.7](/18-optimisation-performance/07-compactage-automatique.md)) ;
- **architecture front-end / back-end** soignée, avec une copie locale du front-end par poste ([section 15.1](01-architecture-front-end-back-end.md)) ;
- **limiter les états et requêtes lourds** exécutés simultanément, et optimiser les requêtes ([chapitre 18](/18-optimisation-performance/README.md)).

## Reconnaître qu'on a dépassé ACE

Certains signaux indiquent qu'une application a dépassé ce que le moteur fichier peut offrir : un nombre d'utilisateurs en écriture qui dépasse durablement 25 à 50, un back-end qui approche 2 Go, des corruptions répétées liées au réseau, un besoin d'accès distant ou nomade, l'exigence de transactions et d'isolation robustes, ou des contraintes de sécurité et de conformité.

La réponse de fond consiste alors à **migrer le back-end vers un serveur de base de données** — typiquement SQL Server — tout en **conservant Access comme front-end** via des tables liées et des requêtes pass-through. Cette **architecture hybride** lève l'essentiel des limites de concurrence, de taille et de corruption sans sacrifier l'investissement réalisé dans l'interface Access. Elle est étudiée au [chapitre 23](/23-migration-interoperabilite/README.md), notamment aux sections [23.1](/23-migration-interoperabilite/01-pourquoi-migrer-sql-server.md) (pourquoi migrer) et [23.4](/23-migration-interoperabilite/04-architecture-hybride-access-sql-server.md) (architecture hybride). Lorsque le besoin est web ou mobile, une migration vers une plateforme cloud (SharePoint/Dataverse, Power Apps) peut être plus pertinente.

## Points de vigilance

- **Le plafond de 255 est théorique** : 25 à 50 utilisateurs actifs en écriture est la limite pratique réaliste.
- **Le réseau est le maillon faible** : interruption pendant une écriture = principale cause de corruption ; filaire obligatoire, Wi-Fi et VPN proscrits pour le back-end.
- **2 Go** pour le back-end : surveiller la taille, archiver à temps (18.9).
- **`MaxLocksPerFile` (9500) et l'erreur 3052** : préférer le fractionnement des grosses opérations à l'augmentation de la limite.
- **Pas de serveur, pas d'isolation, pas de web natif** : des contraintes inhérentes, pas des bugs.
- **Migrer le back-end, garder le front-end** : la voie de sortie naturelle quand les limites sont atteintes (chapitre 23).

## En résumé

Le moteur ACE étant un moteur **fichier** et non client-serveur, tout le traitement se fait côté client et toutes les données transitent par le réseau. Il en résulte des limites concrètes : un maximum de **255 utilisateurs** mais une recommandation pratique de **25 à 50** en écriture intensive ; une taille de base plafonnée à **2 Go** ; une forte **dépendance au réseau**, dont les interruptions sont la première cause de corruption — d'où l'exigence d'une connexion filaire et le rejet du Wi-Fi et du VPN ; une limite de **9 500 verrous par fichier** par défaut (`MaxLocksPerFile`), au-delà de laquelle survient l'erreur 3052 ; un verrouillage **coopératif** sans gestionnaire central ni niveaux d'isolation ; et l'absence d'accès web natif. Ces limites se repoussent par de bonnes pratiques (réseau filaire, transactions courtes, verrouillage au niveau enregistrement, compactage, fractionnement des opérations), mais lorsqu'elles sont durablement atteintes, la solution est la **migration du back-end vers un serveur**, en conservant Access comme front-end.

---

## Conclusion du chapitre 15

Ce chapitre a couvert l'ensemble des questions soulevées par l'usage **partagé** d'une application Access. Il est parti de l'**architecture** front-end / back-end ([15.1](01-architecture-front-end-back-end.md)) et des **modes de partage** ([15.2](02-modes-partage.md)), fondations de tout déploiement multi-utilisateur. Il a ensuite traité le cœur du sujet — le **verrouillage** : ses [stratégies](03-strategies-verrouillage.md) optimiste et pessimiste, la [gestion des conflits de verrou](04-detection-conflits-verrouillage.md), et la propriété [`RecordLocks`](06-recordlocks-formulaires.md) qui l'exprime au niveau des formulaires. Il a abordé des problèmes concrets de la concurrence — la [numérotation séquentielle fiable](05-numeros-sequentiels-multi-utilisateurs.md) — et les outils de gestion de l'état et des sessions : les [TempVars](07-tempvars.md) et le [suivi des sessions utilisateurs](08-gestion-sessions-utilisateurs.md). Il a montré comment [lier le back-end par code](09-tables-liees-par-code.md) pour rendre l'application portable, avant de dresser, dans cette dernière section, les [limites du moteur ACE](10-limites-moteur-ace.md).

Une idée centrale traverse l'ensemble : **concevoir pour le multi-utilisateur dès le départ**. Les problèmes de concurrence sont intermittents et difficiles à diagnostiquer ; ils ne se corrigent pas après coup aussi facilement qu'ils s'évitent par une bonne conception — séparation front-end / back-end, verrouillage optimiste accompagné d'une gestion des conflits, transactions courtes, numérotation protégée, réseau fiable.

Ce chapitre prolongeait directement le [chapitre 14](/14-transactions/README.md) : transactions et verrouillage sont les deux faces de la fiabilité en environnement partagé. Il appelle à son tour deux suites naturelles : le [chapitre 18](/18-optimisation-performance/README.md), qui approfondit la **performance** d'une base partagée (réseau, compactage, limite de taille), et le [chapitre 23](/23-migration-interoperabilite/README.md), qui apporte la réponse de fond — la **migration vers un serveur** — lorsque le moteur fichier atteint ses limites.

⏭️ [16. Programmation orientée objet en VBA Access](/16-poo-vba-access/README.md)
