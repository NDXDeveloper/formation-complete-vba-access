🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 14.4. Transactions imbriquées

Les sections 14.2 et 14.3 ont renvoyé à plusieurs reprises vers ce point particulier : l'**imbrication** des transactions. Il s'agit d'ouvrir une transaction alors qu'une autre est déjà en cours, créant ainsi plusieurs niveaux emboîtés. C'est un mécanisme parfois utile, mais dont la sémantique surprend, et dont la combinaison avec la gestion d'erreur demande une réelle rigueur. Cette section en détaille le fonctionnement, les règles à respecter et les limites propres au moteur ACE.

## Le principe de l'imbrication

Une **transaction imbriquée** est une transaction ouverte à l'intérieur d'une transaction déjà active. On parle alors de niveaux : la première transaction ouverte est le niveau externe (ou de niveau supérieur), celles ouvertes ensuite sont des niveaux internes.

Concrètement, chaque appel à `BeginTrans` non encore résolu ajoute un niveau, et chaque `CommitTrans` ou `Rollback`/`RollbackTrans` résout le niveau le plus interne encore ouvert. En ADO, on l'a vu, `BeginTrans` retourne précisément le numéro du niveau atteint (1, 2, 3…), ce qui matérialise cette notion de profondeur.

## Deux règles fondamentales

Le comportement de l'imbrication tient en deux règles qu'il faut absolument intégrer.

**Première règle — on résout de l'intérieur vers l'extérieur.** Lorsque des transactions sont imbriquées, il faut résoudre la transaction courante avant de pouvoir résoudre une transaction d'un niveau supérieur. Autrement dit, on ne peut pas valider ou annuler directement le niveau externe tant qu'un niveau interne reste ouvert : chaque niveau doit être conclu dans l'ordre inverse de son ouverture.

**Seconde règle — rien n'est réellement écrit avant la validation du niveau le plus externe.** C'est le point le plus contre-intuitif. L'exécution d'un CommitTrans sur une transaction interne ne libère pas les verrous et n'écrit pas réellement les données sur le disque : ce n'est qu'à la validation de la transaction la plus externe que les modifications sont effectivement enregistrées. Un `CommitTrans` interne ne « confirme » donc pas définitivement : il fait simplement passer le travail au niveau englobant, qui garde le pouvoir de décision final.

La conséquence directe est essentielle : une fois un CommitTrans effectué, les changements ne peuvent plus être annulés — sauf si cette transaction est elle-même imbriquée dans une transaction de niveau supérieur qui, elle, fait l'objet d'un rollback. En clair, **un `Rollback` du niveau externe annule tout ce qu'il contient, y compris les transactions internes pourtant validées**. Une validation interne n'est jamais à l'abri d'une annulation externe.

C'est ce qui distingue l'imbrication Jet/ACE des véritables points de reprise (*savepoints*) de certains SGBD serveur : ici, valider un niveau interne ne garantit pas sa persistance, et l'on ne peut pas annuler sélectivement un niveau interne tout en conservant le reste du niveau externe au-delà des règles décrites.

## Imbrication en DAO

En DAO, l'imbrication s'opère naturellement sur le `Workspace` en appelant `BeginTrans` plusieurs fois. Deux caractéristiques sont à retenir.

D'abord, le moteur impose une limite de profondeur : il est possible d'imbriquer les transactions dans les bases Jet/ACE jusqu'à cinq niveaux de profondeur. Au-delà, une erreur est levée. En pratique, on n'atteint quasiment jamais cette limite : si un traitement nécessite plus de quelques niveaux, c'est généralement le signe d'une conception à revoir.

Ensuite, la portée reste globale au workspace, comme pour une transaction simple : tous les niveaux s'appliquent à l'ensemble des bases et recordsets ouverts dans cet espace de travail.

## Imbrication en ADO

ADO prend également en charge l'imbrication, la valeur retournée par `BeginTrans` indiquant le niveau courant. La sémantique générale est la même que celle décrite ci-dessus — résolution du plus interne au plus externe, persistance au dénouement du niveau le plus externe. En revanche, le support effectif et la profondeur autorisée **dépendent du fournisseur OLE DB** utilisé : le fournisseur ACE expose la capacité d'imbrication du moteur, mais un fournisseur tiers peut se comporter différemment. Comme toujours en ADO, il convient de se référer aux capacités du fournisseur concerné (voir [section 14.3](03-transactions-ado.md) et [section 10.9](/10-ado-access/09-connexion-bases-externes.md)).

## La restriction ODBC

Un point essentiel, valable pour DAO comme indirectement pour les architectures connectées : **on ne peut pas imbriquer de transactions sur des sources de données ODBC accédées via le moteur ACE**. L'imbrication des transactions n'est pas possible lorsqu'on accède à des sources de données ODBC par l'intermédiaire du moteur de base de données Access. Cette limite concerne directement les applications utilisant des tables liées à un serveur via ODBC : la gestion transactionnelle multi-niveaux doit alors être pensée côté serveur, sujet abordé à la [section 14.7](07-niveaux-isolation.md).

## L'imbrication et la gestion d'erreur

C'est le piège pratique majeur. Reprenons la première règle : un `Rollback` ne résout que le niveau **le plus interne** encore ouvert. Or, lorsqu'une erreur survient au sein d'un niveau interne, plusieurs niveaux peuvent être ouverts simultanément au moment où le gestionnaire d'erreur est atteint. Un unique `Rollback` ne suffirait donc pas : il annulerait le niveau interne mais **laisserait le niveau externe ouvert** — précisément la situation à éviter, puisqu'une transaction laissée ouverte maintient des verrous.

La parade consiste à **suivre le nombre de niveaux ouverts** et à les annuler tous, un par un, dans le gestionnaire d'erreur. L'exemple ci-dessous illustre ce patron en DAO sur un transfert entre deux comptes (cas d'école de l'atomicité). Le compteur `niveaux` tient le décompte des transactions effectivement ouvertes.

```vba
Public Sub TransfertAvecImbrication()
    Dim ws As DAO.Workspace
    Dim db As DAO.Database
    Dim niveaux As Integer            ' nombre de transactions actuellement ouvertes

    On Error GoTo Gestion_Erreur

    Set ws = DBEngine.Workspaces(0)
    Set db = CurrentDb

    ws.BeginTrans: niveaux = niveaux + 1      ' Niveau 1 (externe)

    ' --- Première unité de travail, dans une transaction interne ---
    ws.BeginTrans: niveaux = niveaux + 1      ' Niveau 2
    db.Execute "UPDATE Comptes SET Solde = Solde - 100 WHERE Id = 1;", dbFailOnError
    ws.CommitTrans: niveaux = niveaux - 1     ' Résout le niveau 2 (PAS encore persisté)

    ' --- Seconde unité de travail, dans une autre transaction interne ---
    ws.BeginTrans: niveaux = niveaux + 1      ' Niveau 2 à nouveau
    db.Execute "UPDATE Comptes SET Solde = Solde + 100 WHERE Id = 2;", dbFailOnError
    ws.CommitTrans: niveaux = niveaux - 1     ' Résout ce niveau 2

    ws.CommitTrans: niveaux = niveaux - 1     ' Résout le niveau 1 : écriture effective sur disque

Nettoyage:
    Set db = Nothing
    Set ws = Nothing
    Exit Sub

Gestion_Erreur:
    ' On annule TOUS les niveaux encore ouverts, du plus interne au plus externe
    Do While niveaux > 0
        ws.Rollback
        niveaux = niveaux - 1
    Loop
    MsgBox "Transfert annulé : " & Err.Description, vbExclamation
    Resume Nettoyage
End Sub
```

Si une erreur survient lors du second `UPDATE`, le gestionnaire trouve deux niveaux ouverts (le niveau 1 externe et le niveau 2 interne en cours) et les annule tous les deux via la boucle. Le crédit du premier compte, bien que validé par un `CommitTrans` interne, est lui aussi défait — illustration concrète de la seconde règle : **le `Rollback` du niveau externe efface tout**.

## À quoi sert réellement l'imbrication ?

Compte tenu de ces contraintes, l'imbrication n'apporte un bénéfice que dans des cas précis :

- **la composition de procédures** : une procédure réutilisable qui gère sa propre transaction (`BeginTrans`…`CommitTrans`) fonctionne aussi bien appelée seule qu'appelée à l'intérieur d'un traitement plus large. Dans le second cas, sa validation ne résout que son niveau interne, et c'est le traitement englobant qui décide de la persistance finale. L'imbrication assure ainsi que des unités de code modulaires se combinent correctement ;
- **le regroupement d'unités de travail** sous une décision globale, lorsque l'on veut pouvoir tout annuler en bloc même si chaque sous-traitement a, localement, « validé ».

En dehors de ces besoins, **une transaction unique et plate est presque toujours préférable**. Elle est plus simple à raisonner, plus simple à annuler (un seul niveau), et évite tout le risque lié au décompte des niveaux dans la gestion d'erreur. L'imbrication doit donc rester un outil réservé aux situations qui la justifient réellement.

## Points de vigilance

- **Toujours équilibrer les appels.** Chaque `BeginTrans` doit avoir son `CommitTrans` ou son annulation correspondante. Un déséquilibre laisse un niveau ouvert et maintient des verrous jusqu'à la fermeture de la base.
- **Un `CommitTrans` interne ne persiste rien.** Ne jamais considérer une validation de niveau interne comme définitive : seul le niveau externe écrit sur disque.
- **Un `Rollback` externe annule les validations internes.** En conséquence directe de la règle précédente.
- **Annuler tous les niveaux en cas d'erreur.** Un seul `Rollback` ne résout que le niveau courant ; suivre le décompte des niveaux est indispensable dans le code imbriqué.
- **Pas d'imbrication sur sources ODBC** via le moteur ACE.
- **Verrous maintenus jusqu'au commit externe.** Plus l'imbrication est profonde et longue, plus les verrous sont conservés, au détriment de la concurrence (voir [chapitre 15](/15-multi-utilisateurs/README.md)).
- **Journal de transaction et espace TEMP.** En espace de travail Access, les transactions sont journalisées dans un fichier situé dans le répertoire désigné par la variable d'environnement TEMP ; si ce journal épuise l'espace disponible, le moteur déclenche une erreur d'exécution. Des transactions très volumineuses maintenues ouvertes (cas favorisé par l'imbrication) accroissent ce risque.

## En résumé

Une transaction imbriquée est une transaction ouverte à l'intérieur d'une autre, créant des niveaux que l'on doit résoudre du plus interne au plus externe. La sémantique du moteur ACE repose sur deux règles : un `CommitTrans` interne ne rend rien persistant — seule la validation du niveau le plus externe écrit réellement les données —, et un `Rollback` du niveau externe annule tout ce qu'il contient, y compris les niveaux internes pourtant validés. DAO autorise jusqu'à cinq niveaux d'imbrication ; ADO en dépend du fournisseur ; et aucune imbrication n'est possible sur des sources ODBC via le moteur ACE. La principale difficulté pratique réside dans la gestion d'erreur, qui doit annuler tous les niveaux ouverts. Faute d'un besoin précis de composition ou de regroupement, une transaction unique reste le choix le plus sûr.

La section suivante quitte le mécanisme transactionnel proprement dit pour aborder un problème connexe en environnement partagé : la [gestion des conflits de mise à jour](05-conflits-mise-a-jour.md).

⏭️ [14.5. Gestion des conflits de mise à jour](/14-transactions/05-conflits-mise-a-jour.md)
