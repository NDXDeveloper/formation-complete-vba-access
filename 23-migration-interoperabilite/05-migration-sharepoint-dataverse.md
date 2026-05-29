🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 23.5. Migration vers SharePoint Lists et Dataverse

La section 23.4 a présenté l'architecture hybride avec SQL Server, qui conserve Access tout en confiant les données à un serveur. Mais elle se heurtait à une limite : on reste dans un environnement de bureau, sans accès web ni mobile natif. Pour les organisations qui veulent au contraire ouvrir leurs données au navigateur, à la mobilité et à l'écosystème Microsoft 365, deux destinations cloud se présentent : **SharePoint Lists** et **Microsoft Dataverse**. Cette section les compare, explique comment Access s'y connecte, et surtout aide à choisir, car ces deux plateformes diffèrent énormément par leurs capacités et leur coût.

## Un avertissement préalable : les Access Web Apps n'existent plus

Beaucoup de tutoriels anciens décrivent la publication d'une base Access *sous forme d'application web* sur SharePoint, via une technologie appelée Access Services (ou Access Web Apps). **Cette voie n'existe plus.** Microsoft a cessé la création de ces applications web en 2017 et a définitivement arrêté les applications existantes en 2018. La recommandation officielle de remplacement est désormais Power Apps (associé à Power Automate) pour les solutions web et mobiles.

Il faut donc écarter d'emblée toute idée de « publier une appli Access sur SharePoint ». En revanche, Access Desktop (le format `.accdb`) n'est aucunement concerné par cet arrêt et continue d'être développé. Aujourd'hui, « porter des données Access vers le cloud » signifie concrètement deux choses : **lier des listes SharePoint**, ou **migrer vers Dataverse** (avec, en option, une interface Power Apps). C'est précisément l'objet des deux parties suivantes.

## Deux philosophies de migration cloud

Avant d'entrer dans le détail, deux idées doivent rester présentes à l'esprit.

D'abord, SharePoint Lists et Dataverse ne jouent pas dans la même catégorie. Les listes SharePoint sont un outil de **gestion de listes** dérivé d'une plateforme de collaboration, peu coûteux mais limité comme base de données. Dataverse est une véritable **plateforme de données** relationnelle, beaucoup plus capable, mais payante. Le choix entre les deux dépend autant de la nature des données que du budget.

Ensuite, dans les deux cas se repose la question du **front-end**. On peut conserver l'interface Access (formulaires, états, VBA) et la relier aux données cloud, dans un esprit hybride proche de celui de la section 23.4 ; ou bien reconstruire entièrement l'interface dans Power Apps. Ce dernier choix est lourd de conséquences : **le code VBA ne s'exécute pas sur le web**. Quitter complètement Access pour une application Power Apps implique donc de réécrire toute la logique applicative avec d'autres outils (Power Fx, Power Automate), et d'abandonner l'investissement VBA pour l'interface.

## SharePoint Lists

### Ce que c'est

Une liste SharePoint est une table hébergée dans SharePoint Online, accessible depuis un navigateur, avec partage, gestion des permissions et historique des versions intégrés. Son grand avantage pratique est d'être **incluse dans la plupart des abonnements Microsoft 365**, sans surcoût.

### Comment Access s'y connecte

Access sait **lier des listes SharePoint** comme il lie des tables : les listes apparaissent dans le front-end comme des tables ordinaires, et les formulaires, états et code VBA continuent de fonctionner. On peut également exporter des tables Access vers des listes. Access offre par ailleurs un fonctionnement hors ligne, en mettant en cache les données de liste pour les synchroniser ensuite.

### Les limites à connaître

Les listes SharePoint comportent des limites structurelles importantes. La plus célèbre est le **seuil d'affichage de 5 000 éléments** : au-delà de ce volume, les opérations qui parcourent la liste sont bridées, sauf à s'appuyer sur des colonnes indexées. Une liste peut techniquement contenir bien plus d'éléments, mais les traitements deviennent contraints.

Plus fondamentalement, une liste n'est **pas un moteur relationnel**. Les colonnes de recherche (lookups) existent mais ne constituent pas une véritable intégrité référentielle avec clés étrangères et cascades ; il n'y a pas de jointures exécutées côté serveur, pas de transactions, des capacités de requête réduites et une gestion de la concurrence sommaire. Les performances restent correctes pour des données modestes et de structure plate, mais se dégradent vite sur des charges relationnelles lourdes.

### Quand SharePoint Lists convient

Les listes SharePoint sont pertinentes pour des données peu volumineuses et relativement plates, des scénarios de collaboration, un besoin d'accès navigateur ou mobile au sein de l'organisation, et du reporting léger — bref, des données dont la nature est intrinsèquement « liste ». Elles ne conviennent pas à des applications relationnelles complexes.

## Dataverse

### Ce que c'est

Dataverse est la plateforme de données qui sous-tend Power Apps, Power Automate et Dynamics 365. C'est un véritable service de données cloud : tables (appelées *entités*), relations réelles, types de données riches (y compris image et fichier), **sécurité par rôles**, logique métier, reporting et fonctionnement hors ligne robuste. Il existe deux déclinaisons : **Dataverse** (la version complète) et **Dataverse for Teams** (allégée, intégrée à Microsoft Teams, qu'on peut faire évoluer ultérieurement vers la version complète).

### L'outil de migration intégré à Access

Depuis 2022, Access embarque un **outil de migration et un connecteur Dataverse** : la commande se trouve dans Exporter → Dataverse. Il migre les tables, les relations et les données en quelques minutes. Cette fonctionnalité requiert une version récente de Microsoft 365 (canal actuel/mensuel) et une licence donnant accès à Dataverse. La migration peut être empaquetée dans une *solution* Dataverse.

Le modèle le plus intéressant est le suivant : on peut **conserver le front-end Access** (formulaires, états, requêtes, macros) en le reliant par tables liées aux données désormais hébergées dans Dataverse, tout en construisant **simultanément** des applications Power Apps, des flux Power Automate ou des tableaux de bord Power BI sur ces mêmes données. Dataverse peut donc servir de back-end hybride à la manière de SQL Server, ou de tremplin vers une refonte complète en Power Apps.

### Points d'attention

Le premier point est le **coût** : à la différence des listes SharePoint incluses dans Microsoft 365, Dataverse suppose une **licence premium** (Power Apps ou Dynamics 365), avec un coût par utilisateur qu'il faut intégrer à la décision. La tarification et les capacités évoluant régulièrement, il convient de vérifier les conditions à jour dans la documentation Microsoft.

Sur le plan des données, deux pièges sont à anticiper : les **champs multivalués** d'Access doivent être convertis (en champ de choix) avant migration, et le type à virgule flottante de Dataverse a une plage de valeurs plus étroite que celui d'Access. Enfin, construire une interface web ou mobile passe par Power Apps et Power Automate, c'est-à-dire un **changement de paradigme** et une courbe d'apprentissage : on quitte le modèle des formulaires Access et du VBA.

## Choisir : SharePoint Lists, Dataverse ou SQL Server

Mises en regard, les trois cibles se distinguent nettement.

| Critère | SQL Server (hybride, 23.4) | SharePoint Lists | Dataverse |
|---|---|---|---|
| Accès web / mobile natif | Non (front-end de bureau) | Oui (navigateur, Power Apps) | Oui (Power Apps) |
| Capacité relationnelle | Forte (vrai SGBD) | Faible (listes, pas d'intégrité réelle) | Bonne (relations, logique métier) |
| Coût / licence | Édition gratuite possible (Express) | Inclus dans Microsoft 365 | Premium (Power Apps / Dynamics) |
| Conserver le front-end Access | Oui (tables liées ODBC) | Oui (listes liées) | Oui (tables liées) ou reconstruire en Power Apps |
| Volumétrie / charge | Élevée | Limitée (seuil de 5 000, charge légère) | Moyenne à élevée (selon licence) |
| Idéal pour | Application critique relationnelle conservant Access | Données « liste », collaboration, partage simple | Modernisation vers Power Platform, web et mobile |

En synthèse : SQL Server reste le meilleur choix lorsqu'on a besoin d'une vraie base relationnelle robuste et que l'interface Access donne satisfaction, sans exigence de web. Les listes SharePoint conviennent à des données simples et collaboratives qu'on veut rendre accessibles facilement, sans surcoût, en acceptant leurs limites. Dataverse s'impose lorsqu'on souhaite moderniser durablement l'application vers la Power Platform, avec accès web et mobile, et qu'on accepte le coût des licences et le changement d'outillage.

## Et le code VBA dans tout cela ?

Le sort du code VBA dépend entièrement du front-end retenu. Si l'on **conserve l'interface Access** reliée à des listes SharePoint ou à Dataverse, le code VBA continue de s'exécuter ; mais ces back-ends ne se comportent pas comme SQL Server. SharePoint, en particulier, n'offre pas de moteur SQL vers lequel envoyer des requêtes Pass-Through, et impose ses seuils et ses limites de requêtage. Le code fonctionne donc, mais avec des contraintes différentes de celles vues à la section 23.3.

Si l'on choisit en revanche de **basculer vers une application Power Apps** accessible par le web, le VBA est purement et simplement abandonné : toute la logique doit être reconstruite avec Power Fx et Power Automate. C'est le coût stratégique majeur de cette voie, et c'est ce qui justifie d'en évaluer sérieusement la faisabilité avant de s'y engager.

## En résumé

Migrer des données Access vers le cloud Microsoft 365 passe aujourd'hui par les listes SharePoint ou par Dataverse — les anciennes Access Web Apps ayant été retirées. Les listes SharePoint sont gratuites et simples mais faibles comme base de données, avec leur seuil de 5 000 éléments. Dataverse est une plateforme bien plus complète, mais premium et porteuse d'un changement de paradigme vers la Power Platform. Dans les deux cas, on peut soit conserver le front-end Access relié aux données cloud, soit reconstruire l'interface en Power Apps — auquel cas le code VBA est abandonné. Cette dernière option, le passage à Power Apps, mérite une analyse de faisabilité à part entière : c'est l'objet de la section suivante.

⏭️ [23.6. Migration vers Power Apps — analyse de faisabilité](/23-migration-interoperabilite/06-migration-power-apps.md)
