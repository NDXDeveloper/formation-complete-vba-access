🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 23.6. Migration vers Power Apps — analyse de faisabilité

La section précédente a montré qu'au moment de porter une application Access vers le cloud, l'une des voies consiste à reconstruire l'interface dans Power Apps — au prix de l'abandon du code VBA. Cette section ne décrit pas *comment* mener cette refonte, mais aide à répondre à la question qui doit la précéder : **est-ce faisable, et est-ce pertinent, pour mon application ?** C'est une analyse de faisabilité, c'est-à-dire un travail d'évaluation lucide, où il s'agit autant d'identifier ce qui ne passera pas que ce qui passera. Mal conduite, une migration vers Power Apps peut aboutir à une application appauvrie ayant coûté des mois de travail ; bien évaluée en amont, elle peut au contraire moderniser durablement une solution.

## L'idée fausse à écarter d'emblée

Le mot « migration » est trompeur dans ce contexte. **Il n'existe aucun outil qui convertit une application Access en application Power Apps.** L'outil de migration vu à la section 23.5 déplace les *données* vers Dataverse, pas l'application : il ne touche ni aux formulaires, ni aux états, ni au code VBA. Passer à Power Apps signifie donc **reconstruire l'application à partir de zéro** sur une plateforme différente, en ne réutilisant que les données une fois celles-ci placées dans une source adaptée.

Cette réalité gouverne toute l'analyse : on n'évalue pas une conversion automatique, mais un **projet de redéveloppement**. La bonne question n'est pas « combien de temps prend la migration ? » mais « combien coûte la reconstruction de cette application, et le résultat sera-t-il à la hauteur de l'existant ? ». Tout le reste de cette section découle de ce constat.

## Ce qu'est Power Apps, pour les besoins de la décision

Power Apps est la plateforme *low-code* de Microsoft pour créer des applications web et mobiles, au sein de la Power Platform. Pour décider, il suffit d'en connaître trois aspects. Il existe deux types d'applications : les **applications canevas** (canvas), où l'on dessine librement l'interface écran par écran et où l'on connecte de nombreuses sources de données, et les **applications pilotées par modèle** (model-driven), générées en grande partie à partir d'un modèle de données Dataverse, plus structurées mais moins libres visuellement. La logique applicative s'écrit en **Power Fx**, un langage de formules proche de celui d'Excel, complété par **Power Automate** pour les traitements et automatisations. Ce sont ces deux outils qui doivent remplacer le code VBA — et Power Fx, il faut le souligner, n'est pas un langage de programmation généraliste : c'est un langage de formules, avec de réelles limites pour la logique complexe.

## Ce qui se transpose, et ce qui ne se transpose pas

L'analyse de faisabilité passe par un inventaire de l'application existante et la projection de chaque élément vers son équivalent Power Apps.

| Élément de l'application Access | Équivalent dans l'écosystème Power Apps | Effort / remarque |
|---|---|---|
| Données (tables) | Source : Dataverse, SharePoint ou SQL Server | Faisable (voir 23.5) |
| Formulaires | Écrans d'application (canvas ou model-driven) | À reconstruire entièrement |
| Requêtes | Filtres Power Fx, vues de la source | À réécrire ; attention à la délégation |
| États (reports) | Power BI, ou Power Automate + modèles Word | Pas d'équivalent direct ; souvent le point le plus dur |
| Code VBA | Power Fx (logique d'écran) + Power Automate (logique serveur) | Réécriture ; Power Fx n'est pas un langage généraliste |
| Macros | Power Fx / Power Automate | Réécriture |
| Automation Office, API Windows, ActiveX | Connecteurs, Power Automate, Office Scripts — ou abandon | Fréquemment bloquant |

Deux lignes de ce tableau méritent une attention particulière, car elles sont les causes d'échec les plus fréquentes.

Les **états** d'Access — avec leurs regroupements, leurs sous-états, leurs totaux et leur mise en page d'impression précise (chapitre 7) — n'ont pas d'équivalent direct dans Power Apps, dont les capacités de reporting natif sont faibles. On les reconstitue généralement avec Power BI (pour l'analyse) ou avec Power Automate associé à des modèles Word (pour les documents imprimables), ce qui constitue souvent un sous-projet à part entière. Une application riche en états sophistiqués est, de ce seul fait, un candidat délicat.

Le **code VBA**, ensuite, ne se traduit pas mécaniquement. La logique d'interface migre vers Power Fx, la logique de traitement vers Power Automate. Le changement de paradigme est profond : on passe d'un langage procédural à un langage de formules réactif.

```
' VBA (Access) : filtrer puis parcourir
Set rs = db.OpenRecordset( _
    "SELECT * FROM Clients WHERE Statut = 'Actif'")
Do Until rs.EOF : '... : rs.MoveNext : Loop
```

```
// Power Fx : une formule déclarative liée à une galerie
Filter(Clients, Statut = "Actif")
```

Pour des règles simples, cette transposition est aisée. Pour des algorithmes VBA volumineux et complexes, elle peut devenir coûteuse, voire impraticable sans recourir massivement à Power Automate.

## Les critères de faisabilité

À partir de cet inventaire, on peut classer l'application.

Sont de **bons candidats** les applications de type « formulaire sur données » (saisie, consultation, flux de validation simples), celles pour lesquelles un accès web ou mobile est réellement nécessaire, celles dont la logique métier reste modérée et le volume de données raisonnable, et celles d'organisations déjà installées dans Microsoft 365 et la Power Platform. Dans ces cas, Power Apps apporte une vraie valeur (mobilité, partage, maintenance centralisée) sans que la reconstruction soit hors de portée.

Sont au contraire des **signaux de difficulté** : une logique VBA lourde et complexe (plusieurs milliers de lignes, algorithmes élaborés), des états d'impression sophistiqués, une dépendance forte à l'automation Office, aux API Windows ou aux contrôles ActiveX (chapitre 22) — qui n'existent tout simplement pas dans Power Apps — de très gros volumes assortis de requêtes relationnelles complexes, des traitements par lots importants, et une intégration étroite avec le poste de travail, les fichiers locaux ou les imprimantes. Plus ces caractéristiques sont présentes, plus la faisabilité décroît et plus le coût de reconstruction grimpe.

## Le piège technique majeur : la délégation

Un concept propre à Power Apps doit absolument entrer dans l'analyse, car il conditionne la faisabilité dès que les données sont volumineuses : la **délégation**. Power Apps cherche à déléguer le traitement (filtres, tris) à la source de données. Lorsqu'une opération ne peut pas être déléguée — parce que la fonction Power Fx employée ou la source ne le permet pas — Power Apps ne traite que les premiers enregistrements rapatriés, dans la limite d'un seuil (500 par défaut, configurable jusqu'à 2 000). Au-delà, les données ne sont tout simplement pas prises en compte, **silencieusement**.

C'est l'équivalent, en plus strict, du piège « ramener tout côté client » vu à la section 23.3. Concrètement, pour une application manipulant beaucoup de données, il faut n'employer que des fonctions *délégables* sur des sources qui supportent la délégation (Dataverse et SQL Server la prennent en charge pour de nombreuses opérations, d'autres connecteurs beaucoup moins). Une application portant sur de gros volumes mais reposant sur des opérations non délégables sera fonctionnellement cassée : c'est un point de faisabilité de premier ordre, à vérifier avant tout engagement.

## Le coût : licences et effort

Deux coûts distincts doivent être chiffrés.

Le **coût de licence** d'abord. Une application Power Apps utilisant des connecteurs premium — notamment Dataverse ou SQL Server — requiert des licences Power Apps premium, facturées par utilisateur ou par application. Les applications limitées aux connecteurs standard (SharePoint, services Microsoft 365) peuvent être couvertes par les licences Microsoft 365 existantes, mais dès qu'on vise une vraie base de données, la facture par utilisateur peut devenir significative à l'échelle de l'organisation. La tarification évoluant régulièrement, il est indispensable de vérifier les conditions à jour auprès de Microsoft.

Le **coût d'effort** ensuite, souvent sous-estimé : la reconstruction de l'application elle-même (de quelques semaines à plusieurs mois selon la complexité), la **formation des développeurs** à Power Fx et Power Automate (un savoir-faire différent du VBA), la **formation des utilisateurs** à une nouvelle interface, la mise en place d'une **solution de reporting** distincte (Power BI) lorsque les états sont importants, et la **gouvernance** (solutions, environnements, cycle de vie applicatif). Ignorer ces coûts annexes est la principale source de mauvaises surprises.

## Une stratégie pour réduire le risque : la coexistence

L'analyse ne débouche pas nécessairement sur un choix binaire « tout migrer ou ne rien faire ». Une approche souvent plus sage consiste à faire **coexister** Access et Power Apps sur les mêmes données. On place les données dans une source partagée (Dataverse ou SQL Server), on **conserve l'application Access de bureau** pour la saisie lourde, les traitements complexes et les états riches, et on construit une **application Power Apps ciblée** uniquement pour la part qui exige réellement le web ou la mobilité — la consultation par des commerciaux sur le terrain, par exemple. Les deux travaillent sur les mêmes données, sans redévelopper l'ensemble.

Cette coexistence permet aussi une migration **incrémentale**, module par module, plutôt qu'un basculement à haut risque. Pour beaucoup d'applications, elle représente le meilleur compromis entre modernisation et préservation de l'existant.

## Conduire l'analyse de faisabilité

En pratique, l'analyse se mène méthodiquement. On dresse d'abord l'**inventaire** de l'application : volume des données, nombre et complexité des formulaires, complexité des états, volume et nature du code VBA, intégrations externes (Office, API, ActiveX, fichiers), besoins hors ligne, nombre d'utilisateurs, et nécessité réelle d'un accès web ou mobile. On **projette** ensuite chaque élément vers son équivalent Power Apps en signalant les écarts, en accordant une attention spéciale aux états, à la logique VBA complexe, aux intégrations bureautiques et à la délégation. On **chiffre** alors l'effort de reconstruction et le coût de licence. On **décide** enfin entre trois options : une migration complète vers Power Apps, une approche hybride de coexistence, ou le maintien d'une architecture Access — par exemple l'hybride SQL Server de la section 23.4 — si le web n'est pas un vrai besoin.

La conclusion la plus fréquente d'une analyse honnête n'est d'ailleurs ni « tout migrer » ni « ne rien faire », mais une forme de coexistence : c'est généralement le signe d'une décision mûrie plutôt que d'un enthousiasme ou d'un rejet de principe.

## En résumé

Migrer vers Power Apps n'est pas une conversion mais une reconstruction : aucun outil ne transforme une application Access en application Power Apps, seules les données se transfèrent. La faisabilité dépend de la nature de l'application — les solutions « formulaire sur données » avec un vrai besoin de mobilité sont de bons candidats, tandis qu'une logique VBA lourde, des états sophistiqués ou des intégrations bureautiques poussées la compromettent. Trois facteurs doivent impérativement entrer dans l'évaluation : la délégation (qui peut casser les applications à gros volume), le coût des licences premium, et l'effort de reconstruction et de formation. La coexistence Access + Power Apps sur des données partagées offre souvent la meilleure issue. Cette analyse close le panorama des migrations vers l'écosystème Microsoft. Reste un dernier cas, celui des organisations qui restent sur Access mais doivent dialoguer avec d'autres moteurs de bases de données non-Microsoft — PostgreSQL, MySQL, MariaDB — objet de la section suivante.

⏭️ [23.7. Cohabitation Access avec d'autres SGBD (PostgreSQL, MySQL, MariaDB)](/23-migration-interoperabilite/07-cohabitation-autres-sgbd.md)
