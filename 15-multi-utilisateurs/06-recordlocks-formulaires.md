🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 15.6. Propriété RecordLocks des formulaires

La [section 15.3](03-strategies-verrouillage.md) a présenté les stratégies optimiste et pessimiste de façon générale. Sur un formulaire lié, c'est la propriété **`RecordLocks`** qui matérialise ce choix : elle détermine comment les enregistrements sont verrouillés lorsque plusieurs utilisateurs modifient les données, et ce qui se produit quand deux d'entre eux tentent d'éditer le même enregistrement. Cette section détaille cette propriété, ses trois valeurs, son paramétrage et ses comportements particuliers.

## Rôle de la propriété

`RecordLocks` est une propriété en **lecture/écriture** qui pilote le verrouillage des enregistrements de la source d'un formulaire en environnement multi-utilisateur. Bien qu'on la rencontre surtout sur les formulaires, elle existe aussi sur les **états** (verrouillage pendant l'aperçu ou l'impression, pour garantir une vue cohérente) et sur les **requêtes** sauvegardées (typiquement une requête action, pendant son exécution). Le principe est le même partout : contrôler quels enregistrements sont verrouillés, et quand.

## Les trois valeurs

La propriété accepte trois valeurs, identiques à celles du réglage global vu à la [section 15.2](02-modes-partage.md) :

- **`0` — Aucun verrou** (*No Locks*) : verrouillage **optimiste**. Aucun verrou n'est posé pendant l'édition ; les utilisateurs sont avertis seulement au moment de l'enregistrement si la donnée a changé entre-temps. C'est le cas à privilégier en mono-utilisateur, ou en multi-utilisateur avec une gestion des conflits.
- **`1` — Tous les enregistrements** (*All Records*) : le plus restrictif. Il sert à garantir qu'aucune modification n'intervient sur les données pendant l'aperçu ou l'impression d'un état, ou pendant l'exécution d'une requête action (ajout, suppression, mise à jour, création de table). Sur un formulaire, il verrouille l'ensemble des enregistrements de la source.
- **`2` — Enregistrement modifié** (*Edited Record*) : verrouillage **pessimiste**. À utiliser lorsqu'on veut empêcher deux utilisateurs ou plus de modifier la même donnée en même temps : l'enregistrement est verrouillé dès le début de la modification.

Attention à l'ordre : les valeurs numériques **ne traduisent pas une sévérité croissante**. La moins restrictive est `0` (Aucun verrou), la plus restrictive est `1` (Tous les enregistrements), et `2` (Enregistrement modifié) se situe entre les deux. Mieux vaut raisonner par libellé que par numéro.

## Lien avec la stratégie de verrouillage

La correspondance avec les notions de la section 15.3 est directe : « Aucun verrou » est l'expression de la stratégie **optimiste**, « Enregistrement modifié » celle de la stratégie **pessimiste**, et « Tous les enregistrements » un verrouillage global réservé aux opérations qui exigent une vue figée.

Il faut garder à l'esprit que la granularité vue en 15.3 reste en jeu : en mode « Enregistrement modifié », le verrou porte sur la **page** contenant l'enregistrement — et peut donc affecter des voisins — sauf si le verrouillage au niveau enregistrement est activé.

## Le défaut et la portée

La valeur par défaut d'un formulaire provient du réglage global « Verrouillage des enregistrements par défaut » (section 15.2) : ce réglage fixe la valeur attribuée aux formulaires **au moment de leur création**. Conséquence importante : modifier ce défaut global n'a **aucun effet sur les formulaires déjà créés**, qui conservent la valeur inscrite dans leur propre propriété `RecordLocks`. Pour changer le comportement d'un formulaire existant, il faut donc modifier sa propriété, et non le réglage global.

## Définir la propriété

On peut renseigner `RecordLocks` de trois manières : par la feuille de propriétés du formulaire (onglet Données), par une macro, ou par code VBA.

```vba
' Lire la valeur (0, 1 ou 2)
Debug.Print Forms!frmClients.RecordLocks

' Définir le verrouillage optimiste
Forms!frmClients.RecordLocks = 0
```

## Comportement à l'exécution : recréation du recordset

Un point essentiel distingue cette propriété : **modifier `RecordLocks` sur un formulaire (ou un état) ouvert provoque la recréation automatique de son recordset**. Autrement dit, changer la valeur en cours d'exécution n'est pas anodin — le jeu d'enregistrements est reconstruit, l'enregistrement courant et la position peuvent être perdus. On préférera donc, dans la grande majorité des cas, fixer la valeur en mode Création plutôt que de la manipuler dynamiquement sur un formulaire actif.

## Le cas ODBC : RecordLocks est ignoré

Lorsque les données proviennent d'une source ODBC (serveur lié), les données d'un formulaire, état ou requête sont traitées **comme si « Aucun verrou » avait été choisi, indépendamment de la valeur de `RecordLocks`**. Ce comportement confirme ce que nous avions vu aux sections [15.3](03-strategies-verrouillage.md) et [14.7](/14-transactions/07-niveaux-isolation.md) : sur une source ODBC, c'est le serveur qui gouverne le verrouillage, et le réglage côté Access reste sans effet. Inutile, donc, de compter sur « Enregistrement modifié » pour un formulaire bâti sur des tables liées à un serveur.

## L'indicateur visuel

En mode Formulaire ou Feuille de données, chaque enregistrement verrouillé affiche un **indicateur de verrou dans son sélecteur d'enregistrement** (à gauche de la ligne). C'est un repère utile pour l'utilisateur, qui voit immédiatement qu'un enregistrement est momentanément indisponible à l'édition.

## Interaction avec l'événement Error

`RecordLocks` détermine *quand* survient le conflit, mais c'est l'événement `Error` du formulaire qui permet de le *traiter* finement, comme évoqué aux sections [14.5](/14-transactions/05-conflits-mise-a-jour.md) et [15.4](04-detection-conflits-verrouillage.md).

En mode « Aucun verrou », une modification concurrente déclenche, à l'enregistrement, la boîte de dialogue de conflit d'écriture intégrée — que l'on peut remplacer par un traitement personnalisé via `Form_Error`. En mode « Enregistrement modifié », c'est un conflit de verrou qui peut survenir lorsqu'un autre utilisateur détient déjà l'enregistrement ; là encore, sur un formulaire lié, l'interception se fait dans `Form_Error`, puisqu'aucun `On Error` n'est actif lors d'une saisie directe dans un contrôle.

## Que choisir ?

Pour la plupart des formulaires d'une application partagée, **« Aucun verrou »** est le bon choix : il maximise la concurrence et, associé à une gestion des conflits (14.5) et à une numérotation fiable (15.5), couvre l'essentiel des besoins. On réserve **« Enregistrement modifié »** aux cas où il est impératif d'empêcher deux modifications simultanées d'une même donnée, en acceptant le blocage qu'il impose aux autres. Quant à **« Tous les enregistrements »**, son usage se limite aux états devant refléter une vue figée et aux requêtes action exigeant qu'aucune donnée ne change pendant leur exécution.

## Points de vigilance

- **Les valeurs ne sont pas dans l'ordre de sévérité** : 0 = Aucun verrou (le moins strict), 2 = Enregistrement modifié, 1 = Tous les enregistrements (le plus strict).
- **Modifier `RecordLocks` sur un formulaire ouvert recrée le recordset** : à éviter en cours d'édition ; privilégier le réglage en mode Création.
- **Le défaut global ne s'applique qu'aux nouveaux formulaires** : les formulaires existants conservent leur propre valeur.
- **`RecordLocks` est sans effet sur une source ODBC** : la donnée est traitée comme « Aucun verrou », le serveur décidant du verrouillage.
- **Le traitement fin des conflits passe par `Form_Error`**, pas par `On Error`, lors d'une saisie directe.
- **En « Enregistrement modifié », le verrou reste au niveau de la page** sans verrouillage au niveau enregistrement (cf. 15.3).

## En résumé

La propriété `RecordLocks` est l'expression, au niveau d'un formulaire (ou d'un état, ou d'une requête), de la stratégie de verrouillage : **0 — Aucun verrou** (optimiste, le choix recommandé par défaut), **2 — Enregistrement modifié** (pessimiste, pour empêcher l'édition simultanée) et **1 — Tous les enregistrements** (le plus restrictif, pour figer les données pendant un état ou une requête action). Sa valeur par défaut découle du réglage global, mais seulement pour les formulaires nouvellement créés. Deux comportements sont à connaître : changer la propriété sur un objet ouvert **recrée son recordset**, et sur une source **ODBC** la propriété est **ignorée** (toujours « Aucun verrou »). Le conflit qu'elle provoque se traite, sur formulaire lié, dans l'événement `Form_Error`.

La section suivante quitte le verrouillage pour un autre outil du multi-utilisateur, dédié à l'état partagé d'une session : les [TempVars, variables globales de session](07-tempvars.md).

⏭️ [15.7. TempVars — variables globales de session](/15-multi-utilisateurs/07-tempvars.md)
