🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 15.3. Stratégies de verrouillage (Optimistic vs Pessimistic)

Le verrouillage est le mécanisme par lequel le moteur arbitre les accès concurrents à une même donnée. Mais « verrouiller » recouvre en réalité **deux questions distinctes** qu'il faut séparer pour raisonner clairement : *quoi* verrouiller — la granularité, page ou enregistrement — et *quand* verrouiller — la stratégie, pessimiste ou optimiste. Cette section traite les deux, leur interaction, et la manière de choisir.

## Deux dimensions à distinguer

La **granularité** détermine l'unité physique verrouillée : une page entière de données, ou un seul enregistrement. C'est une question de *portée* du verrou.

La **stratégie** détermine le moment où le verrou est posé : dès le début de l'édition (pessimiste), ou seulement à l'enregistrement (optimiste). C'est une question de *temporalité* du verrou.

Ces deux dimensions sont indépendantes mais se combinent : on peut avoir un verrouillage pessimiste de page, un verrouillage optimiste d'enregistrement, etc. Les confondre est une source fréquente de malentendus.

## La granularité : page ou enregistrement

Le moteur organise les données en **pages**, des blocs de taille fixe — 2 Ko avant Jet 4, et 4 Ko à partir de Jet 4 (et donc dans ACE). Un enregistrement de longueur variable peut occuper une fraction de page ou s'étendre sur plusieurs.

Avec le **verrouillage de page**, c'est la page entière qui est verrouillée, et non l'enregistrement seul. La conséquence est importante : lorsqu'un utilisateur verrouille un enregistrement, tous les autres enregistrements stockés sur la même page sont **verrouillés collatéralement**, et deviennent inaccessibles aux autres alors même que personne ne les utilise. C'est un effet de bord coûteux en concurrence.

Le **verrouillage au niveau enregistrement**, introduit avec Jet 4 (Access 2000), élimine ces verrous collatéraux : seul l'enregistrement réellement en cours d'usage est verrouillé, tous les autres restent disponibles. Il s'active par l'option « Ouvrir les bases de données avec le verrouillage au niveau de l'enregistrement » (cf. [section 15.2](02-modes-partage.md)).

Plusieurs nuances pratiques doivent toutefois être connues :

- la granularité (page ou enregistrement) est **déterminée par la première session** qui ouvre le fichier ; toutes les connexions suivantes s'y conforment ;
- les **requêtes action utilisent toujours le verrouillage de page**, indépendamment du réglage ;
- le verrouillage au niveau enregistrement **n'est pas toujours honoré** selon la voie d'accès : il ne peut pas être activé via l'ouverture par DAO (`OpenDatabase`), alors que l'interface Access et ADO le permettent. Il arrive donc, en pratique, d'observer plusieurs enregistrements verrouillés malgré l'option cochée.

Autrement dit, le verrouillage au niveau enregistrement réduit la contention mais ne la supprime pas dans tous les cas de figure.

## Le verrouillage pessimiste

Le verrouillage **pessimiste** part du principe que les conflits *vont* se produire, et les prévient en verrouillant la donnée dès le début de l'édition. En DAO, le verrou est posé dès l'appel à la méthode `Edit` (ou `AddNew`) et n'est libéré qu'au `Update` ou à l'annulation. Pendant tout ce temps, aucun autre utilisateur ne peut éditer l'enregistrement (ou la page) concerné.

L'avantage est la **garantie de succès** : une fois le verrou obtenu, la mise à jour est assurée d'aboutir, puisque personne d'autre ne peut modifier la donnée entre-temps. L'inconvénient est la **réduction de la concurrence** : les autres utilisateurs doivent attendre la libération du verrou pour pouvoir agir. Les conflits de verrou — qui obligent à attendre ou font échouer la demande, souvent après un délai d'attente — sont par conséquent plus fréquents avec cette stratégie. Un utilisateur qui ouvre une fiche en édition et s'absente bloque les autres tant qu'il n'a pas validé ou annulé.

La gestion concrète de ces situations de verrou occupé est traitée à la [section 15.4](04-detection-conflits-verrouillage.md).

## Le verrouillage optimiste

Le verrouillage **optimiste** part du principe inverse : les conflits sont *rares*, il est donc inutile de verrouiller pendant l'édition. La donnée n'est verrouillée que brièvement, au moment de l'enregistrement (`Update`). Le reste du temps, plusieurs utilisateurs peuvent éditer librement.

L'avantage est une **concurrence maximale** : les verrous ne sont tenus qu'un court instant. L'inconvénient est qu'on **ne peut pas être certain que l'enregistrement réussira** : si un autre utilisateur a modifié la donnée entre sa lecture et l'enregistrement, une collision survient — l'erreur 3197 en code, ou la boîte de dialogue de conflit d'écriture sur un formulaire lié. Le travail peut alors devoir être repris. La détection et la résolution de ce conflit de mise à jour ont été détaillées à la [section 14.5](/14-transactions/05-conflits-mise-a-jour.md).

## Comment se règle la stratégie

Le choix de la stratégie s'exprime à plusieurs niveaux selon le contexte.

**Sur un recordset DAO**, via la propriété `LockEdits` : `True` pour le pessimiste (c'est la **valeur par défaut** en DAO), `False` pour l'optimiste. On peut aussi la prérégler par l'argument `lockedits` de `OpenRecordset` (`dbPessimistic` pour le pessimiste, les variantes optimistes comme `dbOptimistic` sinon).

**Sur un formulaire lié**, via la propriété `RecordLocks`, qui propose trois valeurs : « Aucun verrou » (optimiste — pas de verrou à l'ouverture de l'enregistrement, conflit vérifié seulement à l'enregistrement), « Enregistrement modifié » (pessimiste — l'enregistrement est verrouillé dès le début de la modification) et « Tous les enregistrements » (le plus restrictif — verrouille l'ensemble des enregistrements de la source du formulaire). Le détail de cette propriété est l'objet de la [section 15.6](06-recordlocks-formulaires.md).

**Globalement**, via le réglage « Verrouillage des enregistrements par défaut » vu à la [section 15.2](02-modes-partage.md).

**Sur un recordset ADO**, via la propriété `LockType`, fixée à l'ouverture : `adLockReadOnly` (lecture seule), `adLockPessimistic` (pessimiste), `adLockOptimistic` (optimiste) ou `adLockBatchOptimistic` (optimiste par lots).

## Illustration en code

Les deux extraits suivants montrent la différence de *temporalité* sur un recordset DAO. La gestion d'erreur — verrou occupé d'un côté, conflit 3197 de l'autre — est volontairement réduite ici ; elle est développée aux sections 15.4 et 14.5.

```vba
' --- PESSIMISTE : verrou dès le Edit ---
Dim rs As DAO.Recordset
Set rs = CurrentDb.OpenRecordset("SELECT * FROM Clients WHERE Id=42;", dbOpenDynaset)
rs.LockEdits = True               ' valeur par défaut en DAO
rs.Edit                           ' verrouille ICI ; erreur immédiate si déjà verrouillé
rs!Solde = rs!Solde + 100
rs.Update                         ' libère le verrou
rs.Close
```

```vba
' --- OPTIMISTE : verrou bref, seulement à l'Update ---
Dim rs As DAO.Recordset
Set rs = CurrentDb.OpenRecordset("SELECT * FROM Clients WHERE Id=42;", dbOpenDynaset)
rs.LockEdits = False              ' optimiste
rs.Edit
rs!Solde = rs!Solde + 100
rs.Update                         ' verrou bref ; erreur 3197 si modifié entre-temps
rs.Close
```

## Une subtilité : l'interaction avec les transactions

Voici un point que beaucoup ignorent et qui relie ce chapitre au [chapitre 14](/14-transactions/README.md). Même en verrouillage optimiste, **une transaction explicite conserve les verrous d'écriture pendant toute sa durée**, ce qui revient de fait à émuler un comportement pessimiste. Croire que l'optimiste reste en vigueur quel que soit le mécanisme transactionnel est une erreur courante.

La conséquence pratique : si l'on encadre des écritures optimistes dans une transaction DAO longue, on récupère les inconvénients de contention du pessimiste, sans en avoir fait le choix. C'est une raison supplémentaire de garder les transactions aussi **courtes** que possible.

## L'interaction avec les tables liées et le serveur

Lorsque les données proviennent d'une source ODBC (serveur lié), le choix optimiste/pessimiste côté Access devient largement théorique : la propriété `LockEdits` d'un recordset DAO est toujours optimiste pour une source ODBC, et c'est le **serveur** qui gouverne réellement le verrouillage et l'isolation. La stratégie de concurrence se règle alors côté serveur, comme exposé à la [section 14.7](/14-transactions/07-niveaux-isolation.md).

## Choisir : quelle stratégie, quand ?

Le choix résulte d'un arbitrage entre garantie et concurrence.

Le **pessimiste** convient lorsque les conflits sont probables *et* les éditions courtes, ou lorsqu'une mise à jour réussie doit être garantie sans possibilité de la refaire (certaines opérations sensibles). Il est à éviter dès que les éditions peuvent durer ou que les utilisateurs risquent de laisser des fiches ouvertes.

L'**optimiste** convient à la grande majorité des applications : forte concurrence, conflits rares, fenêtre d'enregistrement courte, et collisions traitables proprement par la détection (14.5). C'est, en pratique, le choix par défaut recommandé pour la plupart des applications Access partagées — « Aucun verrou » au niveau des formulaires —, à condition d'accompagner ce choix d'une gestion des conflits et d'une numérotation fiable (cf. [section 15.5](05-numeros-sequentiels-multi-utilisateurs.md)).

## Points de vigilance

- **Granularité ≠ stratégie.** Page/enregistrement (quoi) et pessimiste/optimiste (quand) sont deux réglages distincts qui se combinent.
- **Le pessimiste est le défaut DAO.** `LockEdits` vaut `True` par défaut ; à régler explicitement sur `False` pour l'optimiste.
- **Le verrouillage de page verrouille des voisins.** Sans verrouillage au niveau enregistrement, éditer un enregistrement en bloque d'autres sur la même page.
- **Le verrouillage au niveau enregistrement n'est pas garanti partout** : non activable via DAO, ignoré par les requêtes action, fixé par la première session ouvrant le fichier.
- **Optimiste + transaction = pessimiste de fait.** Une transaction tient les verrous d'écriture ; raison de plus pour les garder courtes.
- **Tables liées : c'est le serveur qui décide.** Le choix Access est inopérant sur une source ODBC.

## En résumé

Le verrouillage se règle selon deux axes indépendants : la **granularité** (page — bloc de 4 Ko en ACE, verrouillant les enregistrements voisins — ou enregistrement, qui supprime ces verrous collatéraux mais n'est pas honoré dans tous les cas) et la **stratégie** (pessimiste, qui verrouille dès le début de l'édition et garantit l'enregistrement au prix de la concurrence ; optimiste, qui ne verrouille qu'à l'enregistrement, maximise la concurrence mais expose à un conflit au save). La stratégie se fixe via `LockEdits` en DAO (pessimiste par défaut), la propriété `RecordLocks` des formulaires, le réglage global, ou `LockType` en ADO. Deux subtilités à retenir : une transaction explicite fait tenir les verrous et émule le pessimiste même en mode optimiste, et pour une source ODBC c'est le serveur qui gouverne. Pour la plupart des applications partagées, l'optimiste accompagné d'une bonne gestion des conflits reste le meilleur compromis.

La section suivante se concentre sur le cas pessimiste vécu de l'autre côté : la [détection et la gestion des conflits de verrouillage](04-detection-conflits-verrouillage.md).

⏭️ [15.4. Détection et gestion des conflits de verrouillage](/15-multi-utilisateurs/04-detection-conflits-verrouillage.md)
