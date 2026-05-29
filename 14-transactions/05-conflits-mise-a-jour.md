🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 14.5. Gestion des conflits de mise à jour

Les transactions garantissent l'atomicité d'un traitement, mais elles ne répondent pas à elles seules à une autre question cruciale en environnement partagé : que se passe-t-il lorsque **deux utilisateurs modifient le même enregistrement en même temps** ? Ce cas — le conflit de mise à jour — touche directement à l'intégrité des données, car une gestion naïve conduit à perdre silencieusement les modifications d'un utilisateur. Cette section explique comment ces conflits surviennent, comment les détecter et quelles stratégies adopter pour les résoudre.

> Les **stratégies de verrouillage** (optimiste / pessimiste) sont étudiées en profondeur à la [section 15.3](/15-multi-utilisateurs/03-strategies-verrouillage.md), et les **conflits de verrou** proprement dits (enregistrement inaccessible car déjà verrouillé) à la [section 15.4](/15-multi-utilisateurs/04-detection-conflits-verrouillage.md). La présente section se concentre sur les conflits de **données** : le cas où la modification est possible, mais où la donnée a changé entre-temps.

## Qu'est-ce qu'un conflit de mise à jour ?

Un conflit de mise à jour survient selon le schéma classique « lire — modifier — enregistrer » mené en parallèle par deux sessions :

1. l'utilisateur A lit un enregistrement ;
2. l'utilisateur B lit le même enregistrement ;
3. A modifie sa copie et l'enregistre ;
4. B modifie sa copie (basée sur la version d'origine) et tente de l'enregistrer.

Au moment où B enregistre, la donnée n'est plus celle qu'il avait lue : A l'a déjà changée. Enregistrer la version de B reviendrait à écraser purement et simplement la modification de A, sans que personne ne s'en aperçoive. C'est ce risque de perte silencieuse que la gestion des conflits vise à éviter.

## Le rôle du mode de verrouillage

La manière dont un conflit se manifeste dépend directement du mode de verrouillage retenu. En DAO, ce mode est porté par la propriété `LockEdits` du `Recordset`.

Lorsque `LockEdits` vaut `True` — la valeur par défaut —, le verrouillage est pessimiste : la page contenant l'enregistrement est verrouillée dès l'appel à la méthode `Edit`. Si un autre utilisateur détient déjà cette page, une erreur survient au moment du `Edit`. Dans ce mode, on ne rencontre donc généralement pas de conflit de données à l'enregistrement : on est bloqué plus tôt, à l'édition — ce qui relève des conflits de verrou (section 15.4).

Lorsque `LockEdits` vaut `False`, le verrouillage est optimiste : la page n'est verrouillée qu'au moment du `Update`. Si un autre utilisateur a modifié l'enregistrement entre votre `Edit` et votre `Update`, une erreur survient à l'enregistrement. C'est précisément le conflit de mise à jour qui nous intéresse ici : la modification a été permise, mais elle entre en collision au moment de l'écriture.

On retiendra donc la distinction : **verrouillage optimiste → conflit détecté à l'enregistrement** (objet de cette section) ; **verrouillage pessimiste → blocage dès l'édition** (objet de la section 15.4). À noter que pour les sources de données ODBC accédées via le moteur ACE, `LockEdits` vaut toujours `False` (optimiste), le moteur n'ayant pas la maîtrise des mécanismes de verrouillage des serveurs externes.

Il faut aussi garder à l'esprit que le verrouillage du moteur porte sur la page contenant l'enregistrement, et non sur le seul enregistrement (sauf activation du verrouillage au niveau enregistrement) : un conflit peut donc impliquer des enregistrements voisins situés sur la même page. Ce point est développé au [chapitre 15](/15-multi-utilisateurs/README.md).

## Détecter un conflit

Trois approches se combinent en pratique.

**La détection réactive** consiste à intercepter l'erreur levée par le moteur. En DAO et en verrouillage optimiste, le conflit de mise à jour se traduit par l'**erreur 3197** : le moteur de base de données interrompt l'opération parce que deux utilisateurs tentent de modifier les mêmes données simultanément. Cette erreur survient typiquement lorsque plusieurs utilisateurs ouvrent un enregistrement en verrouillage optimiste et qu'un autre a modifié la donnée entre le `Edit` et le `Update`. C'est l'erreur à capturer pour déclencher une logique de résolution.

**La détection proactive** repose sur une **colonne de version** : un champ (compteur incrémenté à chaque écriture, ou horodatage de dernière modification) dont on relit la valeur juste avant d'enregistrer. Si elle diffère de celle lue initialement, c'est qu'un autre utilisateur est passé entre-temps, et l'on peut traiter le conflit sans même tenter l'écriture. Cette technique est la base de la concurrence optimiste « explicite » et s'avère particulièrement adaptée aux tables liées à un serveur (voir plus bas).

**La relecture et comparaison** consiste à recharger l'enregistrement courant et à comparer ses valeurs à celles que l'on avait lues, pour identifier champ par champ ce qui a changé. C'est la base d'une résolution par fusion.

## Résoudre un conflit : trois stratégies

Une fois le conflit détecté, il faut décider quoi faire. Trois stratégies se dégagent, par ordre croissant de respect des données et de complexité.

- **« Le dernier gagne » (forcer la mise à jour).** On réécrit ses propres valeurs en écrasant celles de l'autre utilisateur. C'est la stratégie la plus simple, mais la plus dangereuse : la modification du premier utilisateur est perdue. À réserver aux cas où cette perte est acceptable ou volontaire.

- **« Le premier gagne » (abandonner sa saisie).** On renonce à sa propre modification, on recharge les données à jour, et l'on laisse l'utilisateur reprendre sa décision sur la base actualisée. En DAO, l'appel à la méthode `Move` avec l'argument `0` rafraîchit l'enregistrement courant avec les modifications faites par l'autre utilisateur — mais cette opération fait perdre vos propres changements en attente. C'est un compromis raisonnable et fréquent.

- **La fusion (présentation à l'utilisateur).** On informe l'utilisateur du conflit, on lui montre sa version et la version actuelle, et on le laisse arbitrer — globalement ou champ par champ. C'est la stratégie la plus respectueuse de l'intégrité des données, mais aussi la plus coûteuse à mettre en œuvre. C'est elle que retiennent les applications soignées sur les données sensibles.

Le choix entre ces stratégies est une décision **métier** autant que technique : il dépend de la valeur des données et du coût d'une perte d'information.

## Exemple : détection et résolution en DAO

L'exemple ci-dessous ouvre un recordset en verrouillage optimiste, tente une modification, et traite le conflit selon la stratégie « le premier gagne ». Les autres stratégies sont indiquées en commentaire.

```vba
Public Sub ModifierTelephone()
    Dim rs As DAO.Recordset

    Set rs = CurrentDb.OpenRecordset( _
        "SELECT * FROM Clients WHERE Id = 42;", dbOpenDynaset)
    rs.LockEdits = False                 ' verrouillage optimiste : conflit détecté à l'Update

    On Error GoTo Conflit

    rs.Edit
    rs!Telephone = "0102030405"
    rs.Update                            ' lève l'erreur 3197 en cas de conflit

    rs.Close
    Set rs = Nothing
    Exit Sub

Conflit:
    If Err.Number = 3197 Then
        ' --- Stratégie « le premier gagne » : on recharge et on abandonne notre saisie ---
        rs.Move 0                        ' rafraîchit avec les données actuelles
                                         ' ATTENTION : nos modifications en attente sont perdues
        MsgBox "Cet enregistrement vient d'être modifié par un autre utilisateur. " & _
               "Vos changements n'ont pas été enregistrés ; les données affichées " & _
               "ont été actualisées. Vérifiez puis ressaisissez si nécessaire.", _
               vbExclamation, "Conflit de mise à jour"
        ' Variantes possibles :
        '   - « le dernier gagne » : relire l'enregistrement, ré-éditer avec nos valeurs, ré-Update
        '   - fusion : comparer nos valeurs et les valeurs actuelles, puis laisser l'utilisateur choisir
    Else
        MsgBox "Erreur " & Err.Number & " : " & Err.Description, vbCritical
    End If

    rs.Close
    Set rs = Nothing
End Sub
```

Le point clé est la double mise en place : `LockEdits = False` rend la détection possible à l'`Update`, et le test `Err.Number = 3197` aiguille vers la résolution choisie.

## Le cas des formulaires liés

Sur un formulaire lié à une table, Access gère lui-même les conflits de mise à jour et affiche, le cas échéant, sa boîte de dialogue « Conflit d'écriture » proposant à l'utilisateur d'enregistrer malgré tout, de copier ses valeurs dans le Presse-papiers, ou d'abandonner ses modifications. Le comportement de verrouillage du formulaire est gouverné par sa propriété `RecordLocks`, traitée à la [section 15.6](/15-multi-utilisateurs/06-recordlocks-formulaires.md).

Pour remplacer le dialogue standard par un traitement personnalisé, on intercepte les erreurs de données dans l'événement `Error` du formulaire et l'on y implémente sa propre stratégie de résolution. Cela permet, par exemple, d'imposer la stratégie « le premier gagne » de façon uniforme, ou d'afficher une interface de fusion maison.

## Côté ADO

En ADO, le mode de verrouillage se choisit à l'ouverture du recordset (`adLockOptimistic`, `adLockPessimistic`, `adLockBatchOptimistic`). En verrouillage optimiste, une violation de concurrence à l'`Update` est remontée par le fournisseur sous forme d'erreur, interceptable comme toute erreur ADO (voir [section 10.8](/10-ado-access/08-gestion-erreurs-ado.md)) ; la propriété `Status` de l'enregistrement permet par ailleurs d'identifier la nature du problème.

Pour les mises à jour par lots (`adLockBatchOptimistic` + `UpdateBatch`), les enregistrements en conflit peuvent être isolés en filtrant le recordset sur les enregistrements en conflit, puis examinés individuellement. Cette approche par lots dépasse le cadre de cette section ; elle est mentionnée ici pour signaler que la philosophie de détection reste la même qu'en DAO.

## Tables liées et serveur

Lorsque les tables sont liées à un serveur (SQL Server notamment), la détection des conflits s'appuie idéalement sur une colonne de version dédiée — un champ de type `rowversion`/`timestamp` que le serveur met à jour automatiquement à chaque modification. Access utilise alors ce champ pour détecter de façon fiable qu'un enregistrement a changé depuis sa lecture. L'absence d'une telle colonne (ou la présence de champs aux types mal interprétés, comme certains booléens à `Null`) est une cause fréquente de conflits de mise à jour intempestifs sur les tables liées. Ces aspects sont approfondis dans le [chapitre 23](/23-migration-interoperabilite/README.md), et la question de l'isolation côté serveur à la [section 14.7](07-niveaux-isolation.md).

## Points de vigilance

- **Ne jamais ignorer un conflit.** Capturer l'erreur 3197 pour la « faire disparaître » sans rien faire revient à perdre des données. Toute détection doit déclencher une vraie résolution.
- **`rs.Move 0` détruit la saisie en attente.** C'est l'effet recherché pour « le premier gagne », mais à utiliser en connaissance de cause.
- **Le verrouillage est par page, pas par enregistrement** (par défaut). Un conflit peut concerner un enregistrement voisin partageant la même page.
- **Tables liées : prévoir une colonne de version.** C'est la condition d'une détection fiable côté serveur, et l'absence de cette colonne génère des faux conflits.
- **Le choix de la stratégie est métier.** « Le dernier gagne » est simple mais destructeur ; la fusion est sûre mais coûteuse. La décision doit être consciente, pas subie.
- **Ne pas confondre conflit de données et conflit de verrou.** Le premier (3197, concurrence optimiste) relève de cette section ; le second (enregistrement déjà verrouillé) relève de la section 15.4.

## En résumé

Un conflit de mise à jour survient lorsque deux sessions modifient simultanément le même enregistrement et que l'enregistrement de la seconde écraserait, sans le savoir, la modification de la première. En verrouillage optimiste (`LockEdits = False` en DAO), ce conflit est détecté à l'`Update` sous la forme de l'erreur **3197** ; en verrouillage pessimiste, le blocage intervient plus tôt, à l'édition. La détection peut être réactive (interception de l'erreur), proactive (colonne de version) ou par relecture. Une fois le conflit détecté, trois stratégies de résolution sont possibles — « le dernier gagne », « le premier gagne » (via `rs.Move 0` en DAO), ou la fusion arbitrée par l'utilisateur —, le choix relevant autant du métier que de la technique. Sur les formulaires liés, Access propose un dialogue intégré que l'on peut remplacer via l'événement `Error`, et sur les tables liées à un serveur, une colonne de version reste la clé d'une détection fiable.

La section suivante quitte la concurrence pour revenir à l'intégrité structurelle des données : les [règles d'intégrité référentielle par code](06-integrite-referentielle.md).

⏭️ [14.6. Règles d'intégrité référentielle par code](/14-transactions/06-integrite-referentielle.md)
