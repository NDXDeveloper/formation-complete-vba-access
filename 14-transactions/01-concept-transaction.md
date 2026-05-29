🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 14.1. Concept de transaction — atomicité et rollback

Avant d'écrire la moindre ligne de code transactionnel, il faut comprendre précisément ce qu'est une transaction et le principe qui en fait toute la valeur. Cette section pose les fondements conceptuels du chapitre : la notion d'unité de travail, les propriétés qui caractérisent une transaction, et les deux dénouements possibles que sont la validation et l'annulation.

## Qu'est-ce qu'une transaction ?

Une **transaction** est une séquence d'opérations sur la base de données traitée comme une **unité logique de travail indivisible**. La propriété qui la définit tient en une phrase : soit toutes les opérations aboutissent et deviennent définitives ensemble, soit aucune ne prend effet et la base revient à son état antérieur. Il n'existe pas de résultat partiel qui survive.

Concrètement, regrouper plusieurs écritures dans une transaction revient à dire au moteur de base de données : « considère ces opérations comme un bloc ; tant que je ne t'ai pas confirmé que tout est bon, ne rends rien définitif ; et si je t'annonce qu'un problème est survenu, fais comme si rien ne s'était passé. »

C'est ce raisonnement « tout ou rien » qui distingue un traitement fiable d'un traitement fragile. Sans transaction, chaque instruction d'écriture est autonome : elle réussit ou échoue indépendamment des autres, et ce qui a déjà été écrit reste écrit, même si la suite échoue.

## Le principe « tout ou rien »

Reprenons l'exemple introduit dans la présentation du chapitre : l'enregistrement d'une commande composée d'un en-tête et de plusieurs lignes.

Sans transaction, le déroulement est le suivant. La procédure écrit l'en-tête : validé immédiatement. Elle écrit la première ligne : validée. La deuxième : validée. À la troisième, une erreur survient (produit inexistant, valeur hors limites, contrainte violée). La procédure s'interrompt. Résultat : l'en-tête et deux lignes subsistent dans la base, la troisième et les suivantes manquent. La commande est incomplète, et rien dans les données ne signale cette anomalie.

Avec une transaction, le même enchaînement produit un résultat radicalement différent. Les écritures sont préparées dans le cadre de la transaction mais ne sont pas rendues définitives. Lorsque l'erreur survient à la troisième ligne, on demande l'annulation de la transaction : l'en-tête et les deux lignes déjà écrites disparaissent, et la base se retrouve exactement dans l'état où elle était avant le début du traitement. Aucune commande partielle, aucune incohérence.

Le principe « tout ou rien » garantit ainsi qu'à aucun moment durable la base ne contient un état intermédiaire incohérent.

## Les propriétés ACID

En théorie des bases de données, une transaction est censée respecter quatre propriétés, désignées par l'acronyme **ACID**. Toutes ne pèsent pas du même poids dans le contexte d'Access, mais les connaître aide à raisonner correctement.

- **Atomicité** (*Atomicity*) — c'est la propriété centrale de cette section. La transaction est indivisible : ses opérations forment un tout. Soit elles s'appliquent toutes, soit aucune. Si une seule échoue, l'ensemble est annulé. C'est l'atomicité qui rend possible le principe « tout ou rien ».

- **Cohérence** (*Consistency*) — une transaction fait passer la base d'un état valide à un autre état valide. Au début comme à la fin, toutes les règles d'intégrité sont respectées. Pendant le déroulement, des états transitoires invalides peuvent exister (par exemple un compte débité mais l'autre pas encore crédité), mais ils ne sont jamais rendus définitifs.

- **Isolation** (*Isolation*) — les transactions concurrentes ne doivent pas interférer entre elles : les états intermédiaires d'une transaction en cours ne devraient pas être visibles par les autres utilisateurs. Le moteur ACE offre un contrôle limité de cette propriété, en particulier comparé à un SGBD serveur. Ce point fait l'objet de la [section 14.7](07-niveaux-isolation.md).

- **Durabilité** (*Durability*) — une fois la transaction validée, ses effets sont persistants : ils survivent à une fermeture de l'application ou à une panne. Ce qui a été confirmé est inscrit durablement dans le fichier de base de données.

Dans la pratique courante du développement Access, c'est l'**atomicité** qui justifie l'écrasante majorité des usages des transactions, suivie de la cohérence qu'elle permet de préserver.

## Le cycle de vie d'une transaction

Toute transaction explicite suit le même cycle en trois temps :

1. **Le début** — on déclare l'ouverture d'une transaction. À partir de ce moment, les écritures qui suivent sont rattachées à cette transaction et ne sont pas rendues définitives.
2. **Les opérations** — on exécute les écritures qui composent l'unité de travail (insertions, modifications, suppressions).
3. **Le dénouement** — la transaction se termine de l'une des deux manières suivantes, et seulement l'une des deux :
   - **la validation** (*commit*), qui rend définitives toutes les écritures du bloc ;
   - **l'annulation** (*rollback*), qui efface toutes les écritures du bloc et restaure l'état initial.

Une transaction ne reste jamais « ouverte » indéfiniment : elle doit toujours se conclure par une validation ou une annulation. Une transaction laissée ouverte (par exemple parce que l'objet qui la porte est fermé sans validation) est traitée comme une annulation.

## La validation (commit)

La **validation** est l'acte par lequel on confirme au moteur que toutes les opérations de la transaction se sont déroulées comme prévu et doivent devenir permanentes. Une fois la validation effectuée, les écritures sont inscrites durablement, elles deviennent visibles pour les autres utilisateurs, et il n'est plus possible de les annuler par un rollback : la transaction est close.

On ne valide donc qu'**à la fin**, lorsque l'on a la certitude que l'ensemble du traitement a réussi. Valider trop tôt — au milieu de la séquence — reviendrait à perdre le bénéfice de l'atomicité pour les opérations restantes.

## L'annulation (rollback)

Le **rollback** est l'acte par lequel on demande au moteur d'effacer toutes les écritures effectuées depuis le début de la transaction et de ramener la base à son état initial, comme si la transaction n'avait jamais eu lieu.

L'annulation intervient typiquement dans deux situations :

- **de façon explicite**, lorsque le code détecte qu'une erreur ou une condition métier interdit de poursuivre, et décide donc d'annuler le bloc entier ;
- **de façon implicite**, lorsqu'une transaction ouverte se termine sans validation — par exemple si l'objet qui la porte est libéré ou si la session se ferme alors que la transaction n'a pas été confirmée. Les écritures non validées sont alors abandonnées.

Le rollback est le mécanisme qui donne concrètement vie au principe « tout ou rien » : c'est lui qui garantit qu'un traitement interrompu ne laisse aucune trace partielle.

## Transactions implicites et transactions explicites

Il est essentiel de comprendre le comportement **par défaut** d'Access, c'est-à-dire en l'absence de toute transaction explicite.

Lorsqu'aucune transaction n'est ouverte, chaque opération d'écriture est validée **immédiatement et individuellement** dès qu'elle s'achève. On parle parfois de validation automatique (*auto-commit*) ou de transaction implicite : chaque instruction constitue à elle seule une micro-transaction, aussitôt confirmée. C'est ce comportement qui explique qu'une séquence d'écritures non encadrée laisse des résultats partiels en cas d'erreur en cours de route.

À l'inverse, une **transaction explicite** est celle que le développeur ouvre délibérément pour regrouper plusieurs écritures sous une même unité de travail. C'est l'objet de tout ce chapitre. Encadrer un traitement par une transaction explicite, c'est suspendre la validation automatique pour la remplacer par une validation unique, contrôlée, intervenant seulement lorsque tout le bloc a réussi.

## Illustration du principe

Les mécanismes précis de mise en œuvre sont détaillés dans les sections suivantes — en DAO à la [section 14.2](02-transactions-dao.md), en ADO à la [section 14.3](03-transactions-ado.md). À titre purement illustratif, voici la forme générale d'une transaction explicite en DAO, qui matérialise le cycle décrit plus haut :

```vba
Dim ws As DAO.Workspace
Set ws = DBEngine.Workspaces(0)

ws.BeginTrans                  ' 1. Début de la transaction

' 2. Les écritures de l'unité de travail
'    (écriture de l'en-tête de commande,
'     écriture des lignes de commande, etc.)

ws.CommitTrans                 ' 3. Validation : tout devient définitif
```

Ce squelette ne devient réellement utile qu'associé à la gestion des erreurs, car c'est l'erreur qui doit déclencher l'annulation. La structure type combine donc transaction et interception d'erreur :

```vba
Dim ws As DAO.Workspace
Dim transOuverte As Boolean

On Error GoTo Gestion_Erreur

Set ws = DBEngine.Workspaces(0)
ws.BeginTrans
transOuverte = True

' --- Écritures de l'unité de travail ---
'     Si l'une d'elles échoue, l'exécution
'     saute directement à Gestion_Erreur.

ws.CommitTrans                 ' Tout a réussi : on valide
transOuverte = False

Sortie:
    Exit Sub

Gestion_Erreur:
    If transOuverte Then ws.Rollback   ' Une erreur : on annule l'ensemble
    MsgBox "Traitement annulé : " & Err.Description, vbExclamation
    Resume Sortie
```

Dans cet exemple, l'indicateur `transOuverte` sert à savoir s'il existe une transaction à annuler au moment où l'on atteint le gestionnaire d'erreur — détail dont l'importance apparaîtra clairement dans les sections de mise en œuvre.

## Rollback et gestion des erreurs : un duo indissociable

Il faut retenir un principe fondamental : **une transaction sans gestion d'erreur associée n'apporte aucune sécurité réelle**. Le rollback n'a de sens que s'il est déclenché au bon moment, c'est-à-dire lorsqu'une erreur ou une condition d'échec survient. Or c'est précisément le rôle du traitement d'erreur étudié au [chapitre 13](/13-gestion-erreurs/README.md) que d'intercepter ces situations et d'orienter le code vers l'annulation.

Transaction et gestion d'erreur fonctionnent donc toujours ensemble : la première définit le périmètre de ce qui doit être annulé en bloc, la seconde décide quand déclencher cette annulation. C'est la raison pour laquelle la maîtrise du chapitre 13 est un prérequis de celui-ci.

## Points de vigilance

Plusieurs idées reçues et pièges entourent la notion de transaction. Mieux vaut les connaître dès maintenant.

- **Une transaction n'est pas une sauvegarde.** Elle protège la cohérence d'un traitement en cours, mais ne remplace en rien une stratégie de sauvegarde de la base. Le rollback ne restaure que les écritures de la transaction courante, pas un état antérieur quelconque de la base.

- **Un rollback ne récupère pas les valeurs de NuméroAuto consommées.** Lorsqu'un enregistrement à clé NuméroAuto est ajouté à l'intérieur d'une transaction qui est ensuite annulée, le numéro qui avait été attribué est définitivement perdu : il n'est pas réutilisé. Une annulation laisse donc des « trous » dans une séquence de NuméroAuto. C'est un comportement normal, mais qui surprend, et qui explique pourquoi ce type de champ ne convient pas comme numérotation métier continue (voir à ce sujet la [section 15.5](/15-multi-utilisateurs/05-numeros-sequentiels-multi-utilisateurs.md)).

- **Toutes les opérations ne sont pas annulables.** Le moteur ACE n'encadre pas de la même manière toutes les opérations : certaines manipulations structurelles (opérations de définition de données) ne participent pas nécessairement à une transaction et ne peuvent donc pas être annulées par un rollback. Les transactions s'appliquent avant tout aux opérations de manipulation de données (insertions, modifications, suppressions d'enregistrements).

- **Une transaction longue a un coût en environnement partagé.** Tant qu'une transaction est ouverte, le moteur maintient des verrous sur les données concernées, ce qui peut gêner les autres utilisateurs. Une transaction doit donc encadrer l'unité de travail strictement nécessaire, et être conclue le plus rapidement possible. L'interaction entre transactions et accès concurrent est approfondie au [chapitre 15](/15-multi-utilisateurs/README.md).

- **Les transactions peuvent être imbriquées**, avec des règles propres qu'il convient de connaître. Ce cas particulier est traité à la [section 14.4](04-transactions-imbriquees.md).

## En résumé

Une transaction est une unité de travail indivisible régie par le principe « tout ou rien ». Elle s'ouvre, exécute un ensemble d'écritures, puis se conclut nécessairement par une validation (*commit*) qui rend tout définitif, ou par une annulation (*rollback*) qui restaure l'état initial. L'atomicité est la propriété centrale qui rend ce mécanisme utile, et le rollback en est l'expression concrète. Par défaut, Access valide chaque écriture immédiatement ; c'est en ouvrant une transaction explicite que l'on regroupe plusieurs écritures sous une validation unique et contrôlée. Enfin, une transaction n'a de réelle valeur que couplée à une gestion d'erreur rigoureuse, qui déclenche l'annulation au moment opportun.

Les sections suivantes passent du concept à la pratique : la mise en œuvre des transactions [en DAO](02-transactions-dao.md), puis [en ADO](03-transactions-ado.md).

⏭️ [14.2. Transactions DAO (BeginTrans, CommitTrans, Rollback)](/14-transactions/02-transactions-dao.md)
