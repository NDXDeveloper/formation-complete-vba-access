🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 18.9. Limites de taille (2 Go) — stratégies de contournement

Toute base Access se heurte tôt ou tard à un plafond infranchissable : **2 Go par fichier**. Cette limite n'est pas un réglage que l'on pourrait relever ; elle est inhérente au format. La connaître, surveiller sa proximité et adopter les bonnes stratégies de contournement est indispensable dès qu'une application manipule des volumes conséquents — au risque, sinon, de voir le fichier devenir instable ou corrompu du jour au lendemain.

## La limite : 2 Go par fichier

Le plafond de 2 Go s'applique à **chaque fichier** `.accdb` (ou `.mdb`), et non à l'application dans son ensemble. C'est un point déterminant : répartir les données sur plusieurs fichiers multiplie d'autant l'espace disponible.

Cette taille compte **tout** ce que contient le fichier — données, index, objets — ainsi que l'espace inutilisé issu du gonflement. L'espace réellement exploitable est donc inférieur à 2 Go : un fichier fortement gonflé peut atteindre la limite avec bien moins de 2 Go de données utiles. À l'inverse, les tables **liées** ne comptent pas dans la taille du fichier qui les référence : leurs données résident ailleurs (back-end, SharePoint, source ODBC).

## Ce qui se passe à l'approche de la limite

Près du plafond, les opérations d'écriture échouent, la base peut devenir instable, et le risque de corruption augmente fortement. Le franchissement survient parfois brutalement — à l'occasion d'un import, d'une requête de création de table, ou d'un gonflement provoqué par une opération volumineuse. Il faut donc agir **avec de la marge**, et non attendre d'être à quelques mégaoctets de la limite : les performances se dégradent d'ailleurs déjà sensiblement sur les très gros fichiers.

## Surveiller la taille par code

La prévention passe par une surveillance proactive : vérifier périodiquement la taille du back-end et déclencher une alerte (ou une action) bien avant le plafond.

```vba
Public Function BackEndProcheLimite(ByVal chemin As String) As Boolean
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(chemin) Then
        ' Seuil d'alerte à ~1,5 Go pour conserver une marge
        BackEndProcheLimite = (fso.GetFile(chemin).Size > 1610612736#)
    End If
End Function
```

Un tel contrôle, exécuté au démarrage, permet d'avertir l'administrateur ou de lancer un compactage ou un archivage avant que la situation ne devienne critique.

> ℹ️ Le `FileSystemObject` est traité au [chapitre 22.10](../22-api-windows-integration-office/10-filesystemobject.md), le compactage déclenché par seuil à la [section 18.7](07-compactage-automatique.md).

## Stratégie 1 — Séparer, puis répartir sur plusieurs back-ends

La première parade est la **séparation** : front-end et back-end dans des fichiers distincts, chacun disposant de ses propres 2 Go. Le front-end (formulaires, requêtes, code) n'approche jamais la limite ; c'est le back-end qui porte les données.

Lorsqu'un seul back-end ne suffit plus, on **répartit les tables sur plusieurs back-ends** — par exemple un fichier pour les données courantes, un autre pour les archives, un troisième pour des tables volumineuses spécifiques. Le front-end se lie à l'ensemble par des tables liées, et chaque back-end bénéficie de son propre plafond de 2 Go.

> ℹ️ L'architecture front-end / back-end est détaillée au [chapitre 15.1](../15-multi-utilisateurs/01-architecture-front-end-back-end.md), la gestion des tables liées par code au [chapitre 15.9](../15-multi-utilisateurs/09-tables-liees-par-code.md).

## Stratégie 2 — Sortir les données binaires du fichier

Rien ne gonfle une base aussi vite que le stockage de fichiers binaires — images, documents, PDF — dans des champs OLE ou Pièce jointe. Ces types de champ inflent le fichier de façon considérable et le mènent rapidement à la limite.

La bonne pratique consiste à **stocker les fichiers sur le système de fichiers** (un dossier réseau) et à ne conserver dans la table que leur **chemin**, dans un simple champ texte. La base reste légère, et l'accès aux fichiers passe par le système de fichiers, bien mieux outillé pour cela.

> ℹ️ Les champs Pièce jointe et multivalués sont traités au [chapitre 9.13](../09-dao-data-access-objects/13-champs-multivalues-pieces-jointes.md).

## Stratégie 3 — Archiver les données anciennes

Déplacer périodiquement les enregistrements historiques vers un back-end d'**archives** — ou les exporter — maintient le back-end de production allégé. Le bénéfice est double : on s'éloigne de la limite de taille, et l'on améliore les performances, puisque les requêtes courantes opèrent sur un volume réduit.

## Stratégie 4 — Compacter et éviter le gonflement

Une part importante de la taille d'un fichier vient souvent du gonflement, et non des données. Le **compactage** régulier récupère cet espace : un fichier proche des 2 Go peut retomber à une fraction de sa taille s'il était surtout gonflé. On évite par ailleurs les pratiques qui gonflent le fichier de données — créer et supprimer à répétition des tables temporaires dans le back-end (à faire dans une base de travail **locale**), enchaîner les requêtes de création de table, ou laisser active la correction automatique des noms.

> ℹ️ Le compactage par code est détaillé à la [section 18.7](07-compactage-automatique.md).

## La réponse de fond : migrer vers un serveur

Lorsque les données dépassent durablement ce qu'Access peut contenir — ou s'en approchent avec une croissance continue —, le contournement atteint ses limites et la vraie solution devient la **migration vers un moteur serveur**. SQL Server Express, gratuit, autorise jusqu'à 10 Go par base ; les éditions complètes lèvent en pratique toute contrainte de taille pour ces usages. Access conserve alors son rôle de front-end, les tables étant liées au serveur via ODBC (architecture hybride). Pour d'autres contextes, SharePoint ou Dataverse peuvent également constituer une cible.

Cette bascule ne résout d'ailleurs pas que la question de la taille : elle répond aussi aux limites du moteur ACE en environnement concurrent ou distant, évoquées à la [section 18.8](08-optimisation-reseau.md).

> ℹ️ La migration vers SQL Server est traitée au [chapitre 23](../23-migration-interoperabilite/README.md) — notamment les raisons de migrer ([23.1](../23-migration-interoperabilite/01-pourquoi-migrer-sql-server.md)), l'architecture hybride par tables liées ODBC ([23.4](../23-migration-interoperabilite/04-architecture-hybride-access-sql-server.md)) et les cibles SharePoint / Dataverse ([23.5](../23-migration-interoperabilite/05-migration-sharepoint-dataverse.md)). Les limites du moteur ACE en concurrence sont exposées au [chapitre 15.10](../15-multi-utilisateurs/10-limites-moteur-ace.md).

## Points clés à retenir

- La limite de 2 Go s'applique par fichier (pas par application) et n'est pas configurable ; elle englobe données, index et gonflement, si bien que l'espace utile est inférieur à 2 Go.
- Agir avec de la marge : surveiller la taille par code et réagir bien avant le plafond.
- Séparer front-end et back-end, puis répartir les tables sur plusieurs back-ends, multiplie l'espace disponible.
- Stocker les fichiers binaires hors de la base (chemin + système de fichiers) évite le principal facteur de gonflement.
- Archiver les données anciennes et compacter régulièrement maintiennent le back-end léger et performant.
- La réponse de fond aux gros volumes — comme aux limites de concurrence — est la migration des données vers un moteur serveur, Access restant le front-end.

---


⏭️ [19. Débogage et tests](/19-debogage-tests/README.md)
