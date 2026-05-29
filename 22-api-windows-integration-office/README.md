🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 22. API Windows et intégration Office

Aussi riche soit-il, Access n'est pas une île. Une application réelle a presque toujours besoin d'interagir avec ce qui l'entoure : connaître l'utilisateur Windows connecté, lancer un programme externe, générer un document Word à partir de données, envoyer un courriel, exporter vers Excel, interroger un service web, lire ou écrire un fichier sur le disque. Aucune de ces opérations ne relève du cœur d'Access ; toutes deviennent possibles dès lors que VBA tend la main vers l'extérieur — vers le système d'exploitation d'une part, vers la suite Office et le web d'autre part.

Ce chapitre rassemble les techniques qui font d'Access un véritable point d'orchestration, capable de piloter d'autres outils et de s'insérer dans des chaînes de traitement plus vastes. Il marque le passage d'une application repliée sur ses propres tables, formulaires et états, à une application ouverte sur son environnement.

## Pourquoi sortir des limites d'Access

Les fonctions intégrées d'Access et de VBA couvrent l'essentiel des besoins courants, mais elles s'arrêtent aux frontières de l'application. Dès qu'un besoin déborde ce périmètre, c'est vers le système et les autres logiciels qu'il faut se tourner.

Récupérer le nom de la machine ou de l'utilisateur, déclencher l'ouverture d'un fichier dans son application par défaut, suspendre brièvement un traitement : autant d'opérations qui exigent de solliciter directement le système d'exploitation. Produire une facture mise en forme dans Word, alimenter un classeur Excel pour l'analyse, expédier un état par courriel via Outlook, bâtir une présentation : autant de tâches qui supposent de **piloter les autres applications Office**. Enfin, dialoguer avec un service web, échanger des données au format JSON : autant de besoins propres aux architectures connectées d'aujourd'hui, que VBA sait satisfaire avec les composants déjà présents sur le poste.

Maîtriser ces techniques, c'est cesser de buter sur les limites d'Access et faire de lui le chef d'orchestre d'un écosystème logiciel.

## Deux paradigmes : API natives et automation COM

Les capacités de ce chapitre reposent sur deux mécanismes de nature différente, qu'il importe de bien distinguer.

Le premier est l'appel d'**API Windows** : la déclaration, au moyen d'instructions `Declare`, de fonctions natives du système d'exploitation, que le code VBA peut ensuite invoquer. C'est une approche de bas niveau, puissante mais exigeante, qui réclame une rigueur particulière dans le typage des paramètres et dans la prise en compte de l'architecture 32/64 bits. Le système ne pardonne pas l'imprécision : une déclaration erronée peut faire planter Access sans avertissement.

Le second est l'**automation COM** (parfois appelée automation OLE) : le pilotage d'une autre application — Excel, Word, Outlook, PowerPoint — à travers son **modèle objet**, exactement comme on manipule en VBA les objets d'Access lui-même. On crée une instance de l'application cible, on agit sur ses objets, puis on libère proprement les ressources. Cette approche, plus structurée et plus lisible que l'appel d'API, suppose en contrepartie que l'application pilotée soit installée sur le poste, et impose une gestion soigneuse du cycle de vie des objets.

Le chapitre mobilise tantôt l'un, tantôt l'autre de ces paradigmes, selon ce qui convient le mieux au besoin traité.

## Ce que couvre ce chapitre

Le chapitre s'organise en quatre temps, du plus proche du système au plus ouvert sur le monde extérieur.

Les premières sections posent les **fondations de l'appel d'API Windows** : la déclaration et l'appel des fonctions natives, avec l'attention requise par le mot-clé `PtrSafe` et la compatibilité 32/64 bits, puis un florilège d'**API courantes** d'usage fréquent (nom d'utilisateur, nom de machine, lancement de programmes, temporisation).

Vient ensuite le cœur de l'intégration Office, avec l'**automation des applications de la suite** : l'export et l'import de données avec Excel, la génération de documents avec Word, l'envoi de courriels et la gestion des contacts avec Outlook, et la création de présentations avec PowerPoint. C'est ici qu'Access devient producteur de livrables bureautiques.

Le troisième temps ouvre l'application sur le **web** : l'appel de services REST depuis VBA, au moyen des composants MSXML2 et WinHttp, et le **parsing du format JSON** sans recourir à une bibliothèque externe — deux compétences qui inscrivent Access dans les échanges de données contemporains.

Enfin, les dernières sections traitent l'**interaction avec le système** : la lecture et l'écriture du registre Windows, et la manipulation du système de fichiers au moyen du `FileSystemObject`. Ces capacités, déjà sollicitées par les mécanismes de déploiement du chapitre 21, trouvent ici leur traitement complet.

## Des précautions particulières

Plus que d'autres, ce chapitre exige une discipline constante, car on y manipule des ressources extérieures à Access qui n'offrent pas le même filet de sécurité.

La **compatibilité 32/64 bits** conditionne toute déclaration d'API : un typage incorrect des pointeurs et des handles conduit à des plantages difficiles à diagnostiquer. Ce sujet, central pour les API, a été traité sous l'angle du déploiement à la section 21.7, à laquelle ce chapitre renvoie constamment.

La **libération des objets d'automation** est une autre exigence incontournable : une instance d'Excel ou de Word créée par code, mais non refermée et non libérée, demeure en mémoire sous forme de processus fantôme, invisible mais consommateur de ressources. Refermer les applications pilotées et réinitialiser les variables objets fait partie intégrante de la technique.

La **dépendance aux applications installées** réduit par ailleurs la portabilité : une application qui pilote Word ne fonctionnera pas sur un poste dépourvu de Word. Cette dépendance doit être anticipée et, si possible, gérée avec souplesse.

Enfin, appeler des API et lancer des programmes externes comporte des **implications de sécurité** : ces opérations s'inscrivent dans le cadre de la sécurité des macros (section 20.5) et appellent la même prudence que tout code disposant d'un accès étendu au système.

## Objectifs d'apprentissage

À l'issue de ce chapitre, vous serez en mesure de :

- Déclarer et appeler des fonctions de l'API Windows depuis VBA, dans le respect de la compatibilité 32/64 bits.
- Exploiter les API les plus courantes pour obtenir des informations système et déclencher des actions externes.
- Piloter Excel, Word, Outlook et PowerPoint par automation, en gérant proprement le cycle de vie des objets.
- Générer des documents, des classeurs, des courriels et des présentations à partir des données d'Access.
- Appeler des services web REST et traiter des données JSON sans bibliothèque externe.
- Lire et écrire le registre Windows, et manipuler le système de fichiers de façon robuste.

## Prérequis

Ce chapitre suppose une bonne maîtrise des fondamentaux de VBA (chapitre 3), en particulier des procédures, des types de données et de la gestion d'erreurs (chapitre 13), indispensable lorsqu'on manipule des ressources externes susceptibles d'échouer. La compréhension des références et du choix entre liaison précoce et liaison tardive (sections 2.5 et 2.6) est directement mobilisée par l'automation Office. Enfin, la familiarité avec la manipulation des modèles objets, acquise au contact de ceux d'Access, de DAO et d'ADO (chapitres 4, 9 et 10), facilite l'abord des modèles objets des autres applications Office.

## Environnement et versions

Les techniques d'automation requièrent que les applications pilotées soient installées sur le poste qui exécute le code. Les déclarations d'API doivent respecter l'architecture 32 ou 64 bits de l'installation (section 21.7). Les sections consacrées aux services web et au JSON s'appuient sur des composants présents en standard sur les versions modernes de Windows, sans installation supplémentaire. Les déclarations d'API portables prêtes à l'emploi sont par ailleurs rassemblées à l'annexe H.

⏭️ [22.1. Déclaration et appel d'API Windows depuis Access (PtrSafe)](/22-api-windows-integration-office/01-declaration-api-windows.md)
