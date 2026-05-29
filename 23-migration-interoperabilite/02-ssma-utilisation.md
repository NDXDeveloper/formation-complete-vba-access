🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 23.2. SQL Server Migration Assistant (SSMA) — utilisation et limites

Une fois la décision de migrer prise (section 23.1), la question devient pratique : comment transférer concrètement les tables et les données d'Access vers SQL Server ? Microsoft propose pour cela un outil gratuit et dédié, **SQL Server Migration Assistant (SSMA)**, qui automatise une grande partie du travail. Cette section présente son rôle, le déroulement d'une migration, puis — c'est au moins aussi important — ce que l'outil ne fait pas. Car SSMA est une aide précieuse, mais il ne réalise qu'une partie du chemin, et croire qu'il prend tout en charge est l'erreur la plus fréquente.

## Qu'est-ce que SSMA ?

SSMA est une famille d'outils Microsoft permettant de migrer différents systèmes de bases de données vers SQL Server et Azure SQL. Il en existe plusieurs déclinaisons (pour Oracle, MySQL, DB2…) ; celle qui nous concerne est **SSMA for Access**. C'est un logiciel gratuit, téléchargeable depuis le site de Microsoft, qui s'installe sur un poste de travail.

SSMA a remplacé l'ancien *Assistant Upsizing* (Upsizing Wizard) qui était intégré directement à Access dans les versions plus anciennes, et qui a été retiré des versions récentes. SSMA est aujourd'hui la voie recommandée par Microsoft pour une migration assistée d'Access vers SQL Server.

L'outil remplit trois fonctions principales : il **convertit le schéma** (tables, types de données, clés, index, relations) d'Access vers son équivalent SQL Server ; il **migre les données** contenues dans les tables ; et, en option, il **relie** ensuite le front-end Access aux tables désormais hébergées sur le serveur, de sorte que l'application continue de fonctionner sans que l'utilisateur perçoive le changement.

## Préparatifs avant de lancer SSMA

Quelques précautions s'imposent avant de démarrer, et les négliger expose à des surprises.

Il faut d'abord travailler sur une **copie** de la base, jamais sur le fichier de production. Cette copie doit au préalable être **compactée et réparée**, et il est judicieux d'en profiter pour nettoyer les données (doublons, valeurs aberrantes, enregistrements orphelins) : ce qui est sale dans Access le restera dans SQL Server, et il est bien plus simple de nettoyer en amont.

Côté logiciel, outre SSMA lui-même, Microsoft fournit un **Extension Pack** qui s'installe sur le serveur SQL. Il ajoute des fonctions et objets auxiliaires servant notamment à émuler certains comportements d'Access et à faciliter la migration des données. Son installation suppose des droits suffisants sur le serveur ; dans certains environnements verrouillés, cette dépendance peut poser problème, point sur lequel nous reviendrons.

Il faut enfin disposer des **droits nécessaires** sur l'instance SQL Server cible (création de base, de schéma, d'objets) et avoir tranché en amont les choix d'architecture : version et édition de SQL Server visées (ou Azure SQL), nom de la base et du schéma cible.

## Le déroulement d'une migration avec SSMA

Le processus suit une séquence d'étapes assez constante :

1. **Créer un projet SSMA**, en indiquant la cible de migration (la version de SQL Server ou Azure SQL visée).
2. **Ajouter la ou les bases Access** au projet ; SSMA en analyse alors le contenu et affiche l'arborescence des objets source.
3. **Se connecter à l'instance SQL Server cible**, en précisant la base de destination.
4. **Générer le rapport d'évaluation** (assessment), qui signale les points de conversion délicats — à lire impérativement, voir plus bas.
5. **Vérifier et ajuster les correspondances de types** proposées entre Access et SQL Server.
6. **Convertir le schéma** : SSMA produit la définition des objets cibles.
7. **Synchroniser avec la base** : les objets sont réellement créés sur le serveur.
8. **Migrer les données** depuis Access vers les tables SQL Server.
9. **Relier les tables** (en option) pour reconnecter le front-end Access au serveur.

Deux étapes méritent qu'on s'y arrête, car elles prêtent souvent à confusion.

La distinction entre **convertir** et **synchroniser** est essentielle. La conversion produit, *en local dans SSMA*, une représentation du schéma cible que l'on peut examiner et corriger avant tout impact sur le serveur. Ce n'est que la synchronisation qui applique réellement ces objets à la base SQL Server. Autrement dit, convertir n'écrit rien sur le serveur ; on peut convertir, inspecter le résultat, ajuster, puis seulement synchroniser une fois satisfait.

L'étape de **liaison des tables** est celle qui rend l'architecture hybride opérationnelle (section 23.4). Lorsqu'elle est activée, SSMA conserve les tables locales d'origine — généralement en les renommant avec un suffixe — et crée à leur place des **tables liées** portant les noms d'origine et pointant, via ODBC, vers le serveur. Comme les noms sont préservés, les formulaires, états et requêtes qui référencent ces tables continuent de fonctionner sans modification de nom. C'est précisément ce qui permet à l'application de tourner sur SQL Server tout en gardant son interface Access.

## Le rapport d'évaluation : à lire avant tout

Avant même de convertir, SSMA peut produire un **rapport d'évaluation** qui examine chaque objet et signale ce qui se convertira proprement, ce qui nécessitera une intervention manuelle, et ce qui ne peut pas être converti. Ce rapport est l'outil le plus utile pour anticiper les difficultés et estimer la charge de travail réelle d'une migration. Le parcourir attentivement avant de lancer quoi que ce soit évite les mauvaises surprises et permet de planifier les corrections nécessaires. Le négliger, c'est découvrir les problèmes une fois la migration faite, au pire moment.

## Les limites et les angles morts de SSMA

C'est ici que se joue la bonne compréhension de l'outil. SSMA fait gagner un temps considérable sur le schéma et les données, mais il comporte des limites qu'il faut connaître précisément.

### Le code VBA n'est pas migré

C'est la limite la plus importante, et celle qui surprend le plus. **SSMA migre le schéma et les données, pas l'application.** Les formulaires, les états, les macros et tout le code VBA ne sont absolument pas pris en charge : ils restent tels quels dans le front-end Access. L'adaptation de ce code aux particularités de SQL Server — qui est souvent le travail le plus délicat de toute la migration — reste entièrement à la charge du développeur. C'est l'objet de la section 23.3. SSMA ne réalise donc, en réalité, qu'une moitié du projet.

### La conversion des requêtes est limitée et peu fiable

SSMA tente de convertir certaines requêtes Access sauvegardées en vues SQL Server, mais le résultat est aléatoire. Le dialecte SQL d'Access (Jet SQL) diffère trop de celui de SQL Server (T-SQL) pour que la traduction soit automatique et fiable. Les **requêtes analyses croisées** (TRANSFORM/PIVOT), les **requêtes paramétrées**, les **requêtes action** et toutes celles qui utilisent des **fonctions propres à Access** échouent fréquemment ou se convertissent de façon incorrecte. En pratique, il faut considérer que les requêtes devront être revues, voire réécrites, manuellement. SSMA donne un point de départ, pas une solution clé en main de ce côté.

### Les pièges de correspondance des types

SSMA propose une correspondance entre les types Access et les types SQL Server, mais celle-ci mérite une vérification systématique. Certaines correspondances sont triviales, d'autres porteuses d'effets de bord à connaître.

| Type Access | Type SQL Server proposé | Point de vigilance |
|---|---|---|
| Texte court | `nvarchar(n)` | Vérifier la longueur retenue |
| Texte long (Mémo) | `nvarchar(max)` | Conversion généralement correcte |
| NuméroAuto | `int IDENTITY` | Le comportement de récupération du nouvel identifiant change en VBA (voir 23.3) |
| Oui/Non | `bit` | Access stocke Vrai à **-1**, SQL Server à **1** : impacte les comparaisons et le code |
| Date/Heure | `datetime2` (ou `datetime`) | Plage et précision différentes de celles d'Access |
| Monétaire | `money` | Fonctionne, mais `decimal` est souvent préférable |
| Numérique décimal | `decimal`/`numeric` | Contrôler précision et échelle |
| Objet OLE | `varbinary(max)` | Souvent à repenser plutôt qu'à transférer tel quel |
| Lien hypertexte | `nvarchar(max)` | Perte de la nature « hyperlien » propre à Access |
| Pièce jointe | *(pas d'équivalent)* | Refonte nécessaire (stockage externe ou table dédiée) |
| Champ multivalué (MVF) | *(pas d'équivalent)* | Refonte en véritable table de liaison |

Le passage du booléen Access (-1) au `bit` SQL Server (1) est un classique : du code qui testait `If champ = -1` ou des requêtes comparant à `-1` ne fonctionneront plus comme attendu. De même, les **champs multivalués et les champs pièce jointe** (vus à la section 9.13) n'ont aucun équivalent natif côté serveur et imposent une refonte du modèle de données, ce que SSMA ne peut décider à votre place.

### Les valeurs par défaut, règles de validation et relations

Les valeurs par défaut et règles de validation exprimées avec des **fonctions propres à Access** ne se traduisent pas toujours. Les **relations et l'intégrité référentielle** peuvent être recréées par SSMA, mais les règles de mise à jour ou de suppression en cascade méritent une vérification au cas par cas après migration.

### Ce que SSMA n'optimise pas

SSMA reproduit la structure existante ; il ne **repense pas** la base pour son nouveau contexte. Les index pertinents pour une charge client-serveur, les ajustements de modèle, l'optimisation des accès distants ne relèvent pas de l'outil. Une base migrée « telle quelle » fonctionnera, mais ne tirera pas spontanément le meilleur parti de SQL Server tant que ce travail n'aura pas été mené.

### Volume, performances de l'outil et environnement

Sur de **grosses bases**, SSMA peut se révéler lent et gourmand en mémoire ; il est alors parfois préférable de migrer par lots plutôt qu'en une seule passe. La dépendance à l'**Extension Pack** côté serveur peut être bloquante dans des environnements très verrouillés. Enfin, SSMA est un outil d'**assistance ponctuelle**, pas un mécanisme de synchronisation continue : il n'est pas conçu pour maintenir Access et SQL Server alignés dans la durée, mais pour réaliser la bascule.

## Quand préférer une migration manuelle

SSMA n'est pas la seule voie. Pour une base **petite ou simple**, ou lorsqu'on souhaite garder un contrôle total sur le schéma cible, une migration **manuelle** est souvent plus propre : on crée le schéma à la main dans SQL Server Management Studio (ce qui permet d'emblée d'adopter de bons types et de bons index), puis on transfère les données à l'aide de l'assistant d'importation/exportation de SQL Server, de requêtes d'ajout via tables liées, ou encore de `DoCmd.TransferDatabase` depuis Access.

SSMA prend tout son sens pour les **schémas volumineux ou complexes** et pour la commodité de la reliaison automatique du front-end. Mais quelle que soit la méthode, ses angles morts — code VBA, requêtes, types particuliers — signifient qu'une part de travail manuel est de toute façon inévitable. Le choix entre SSMA et migration manuelle n'élimine donc pas l'effort d'adaptation ; il en déplace seulement le curseur.

## En résumé

SSMA automatise la conversion du schéma, la migration des données et la reliaison du front-end Access, ce qui représente un gain de temps réel et fait de lui l'outil de référence pour cette étape. Mais il ne traite ni le code VBA, ni de façon fiable les requêtes, et certaines structures de données exigent une refonte qu'aucun outil ne peut décider. La migration des objets et des données n'est donc que la première moitié du projet. La seconde — l'adaptation du code VBA aux différences entre Jet SQL et T-SQL — fait l'objet de la section suivante.

⏭️ [23.3. Adapter le code VBA après migration (T-SQL vs Jet SQL)](/23-migration-interoperabilite/03-adapter-code-apres-migration.md)
