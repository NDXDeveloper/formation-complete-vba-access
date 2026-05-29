🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 18.5. Optimisation des formulaires (RecordSource, filtres, chargement différé)

Les formulaires sont l'un des principaux points sensibles de la performance dans une application Access. À l'ouverture, un formulaire lié charge sa source, peuple ses listes déroulantes, instancie ses sous-formulaires et applique sa mise en forme : autant d'opérations qui, mal maîtrisées, transforment une ouverture en attente perceptible. Le principe directeur tient en une phrase : **charger moins, et plus tard**.

## Un formulaire ne va pas plus vite que sa source

La performance d'un formulaire lié est d'abord celle de son `RecordSource`. Le lier à une table entière ou à un `SELECT *` lui fait récupérer bien plus que nécessaire. On restreint donc les colonnes aux seuls champs affichés ou utilisés, et les lignes à celles dont l'utilisateur a réellement besoin. Comme pour les Recordsets (voir la [section 18.2](02-optimisation-recordsets.md)), une requête sauvegardée — au plan préoptimisé — est préférable à une chaîne SQL recompilée à chaque ouverture, et la source doit s'appuyer sur des index adaptés.

> ℹ️ L'impact de l'indexation est traité à la [section 18.6](06-indexation-performances-sql.md) ; le `RecordSource` dynamique au [chapitre 6.5](../06-formulaires/05-recordsource-filtres-dynamiques.md).

## Limiter la source, au bon endroit

Limiter ce que le formulaire affiche peut se faire à plusieurs niveaux, et l'endroit choisi change le coût.

Le moyen recommandé pour ouvrir un formulaire sur un sous-ensemble est l'argument `WhereCondition` de `OpenForm` : le filtre est pris en compte dès l'ouverture, de sorte que seules les lignes concernées sont chargées.

```vba
DoCmd.OpenForm "frmClients", , , "Ville = 'Paris'"
```

Pour un filtrage **interactif** sur un formulaire déjà ouvert, on emploie les propriétés `Filter` et `FilterOn`, qu'il est facile d'activer ou de désactiver.

```vba
Me.Filter = "Ville = '" & Me.txtVille & "'"
Me.FilterOn = True
```

Dans tous les cas, la priorité est de ne pas charger toute la table à l'ouverture lorsqu'on n'a besoin que d'une fraction des enregistrements. Le filtrage suppose par ailleurs un formatage correct des valeurs (délimiteurs, dates, décimales) et une protection contre l'injection SQL.

> ℹ️ Voir le SQL paramétré ([11.5](../11-sql-access-vba/05-sql-parametree.md)) et la localisation du SQL dynamique ([11.7](../11-sql-access-vba/07-localisation-formatage-sql.md)).

## Le chargement différé : ouvrir vide, charger ensuite

Pour les formulaires de recherche, la technique la plus efficace consiste à ouvrir le formulaire **sans données**, puis à définir sa source une fois les critères saisis. À la conception, on laisse le `RecordSource` vide (ou pointant sur une requête qui ne renvoie rien, par exemple avec une clause toujours fausse) ; l'ouverture est alors instantanée, et l'on n'interroge la base que sur des critères précis.

```vba
' Une fois les critères saisis par l'utilisateur :
Me.RecordSource = "SELECT ClientID, Nom, Ville FROM Clients " & _
                  "WHERE Ville = '" & Me.txtVille & "'"
```

On évite ainsi de charger des milliers de lignes que l'utilisateur n'a jamais demandées, simplement parce qu'il a ouvert l'écran.

## Formulaires de saisie pure : `DataEntry`

Lorsqu'un formulaire ne sert qu'à **ajouter** des enregistrements, la propriété `DataEntry` à `True` l'ouvre positionné pour la saisie sans charger l'existant — l'équivalent, côté formulaire, de l'option `dbAppendOnly` vue pour les Recordsets. L'ouverture devient immédiate quelle que soit la taille de la table.

```vba
Me.DataEntry = True
```

## Listes déroulantes et listes : la source compte aussi

Chaque liste déroulante (`ComboBox`) ou liste (`ListBox`) exécute sa `RowSource` au chargement du formulaire. Une liste alimentée par une table volumineuse ralentit l'ouverture — et reste de toute façon inexploitable pour l'utilisateur. On limite donc sa source en lignes et en colonnes.

Plusieurs approches allègent encore la charge : ne peupler la liste qu'à la demande (au premier déroulement ou à la prise de focus), filtrer une liste dépendante par code en réagissant à la sélection d'une autre (listes en cascade, avec `Requery`), ou s'appuyer sur un cache mémoire pour les petites listes statiques.

> ℹ️ Les contrôles liste sont traités au [chapitre 8.2](../08-controles-evenements/02-controles-texte-listes.md) ; la mise en cache de petites listes relève de la [section 18.4](04-cache-variables-module.md).

## Sous-formulaires : charger à la demande

Chaque sous-formulaire possède sa propre source, qui se charge lors de l'ouverture du formulaire parent. Multiplier les sous-formulaires — en particulier sur les pages d'un contrôle onglet — multiplie d'autant le temps de chargement.

La parade consiste à **différer** le chargement : on n'affecte le `SourceObject` d'un sous-formulaire que lorsque l'utilisateur accède à l'onglet correspondant, plutôt que de tout instancier à l'ouverture.

```vba
Private Sub TabControl0_Change()
    If Me.TabControl0 = 1 Then               ' page « Commandes »
        If Me.sfCommandes.SourceObject = "" Then
            Me.sfCommandes.SourceObject = "sfrmCommandes"
        End If
    End If
End Sub
```

On veille par ailleurs à garder les sources des sous-formulaires épurées, puisque la liaison parent/enfant (`LinkMasterFields` / `LinkChildFields`) les réinterroge à chaque déplacement dans le parent.

> ℹ️ Voir les sous-formulaires ([6.4](../06-formulaires/04-sous-formulaires.md)) et le contrôle onglet ([8.4](../08-controles-evenements/04-controle-onglet.md)).

## Le poison des formulaires continus : les fonctions de domaine dans les contrôles

C'est l'une des erreurs les plus pénalisantes. Placer une fonction de domaine — `=DLookup(...)`, `=DSum(...)`, `=DCount(...)` — dans la source d'un contrôle la fait réévaluer fréquemment (au rafraîchissement, au changement d'enregistrement) ; sur un formulaire **continu** ou en feuille de données, elle est évaluée **pour chaque ligne visible**. L'effet sur les performances est dévastateur.

La bonne pratique est de déplacer le calcul dans la source du formulaire, pour qu'il soit effectué une seule fois par le moteur, via une jointure ou une agrégation :

```text
' Au lieu d'un contrôle =DSum("Montant"; "Lignes"; "CommandeID=" & [CommandeID]) :
SELECT c.CommandeID, c.DateCmd, t.Total
FROM Commandes AS c
LEFT JOIN (
    SELECT CommandeID, SUM(Montant) AS Total
    FROM Lignes GROUP BY CommandeID
) AS t ON c.CommandeID = t.CommandeID
```

> ℹ️ On retrouve ici le raisonnement ensembliste des [sections 18.2](02-optimisation-recordsets.md) et [18.3](03-execute-vs-runsql.md).

## Maîtriser les rafraîchissements

Le rendu visuel a un coût. Lorsqu'un traitement manipule de nombreux contrôles, on suspend le repeint du formulaire le temps de l'opération, ce qui supprime le scintillement et réduit la durée.

```vba
Me.Painting = False
'  ... manipulation de nombreux contrôles ...
Me.Painting = True
```

Il faut aussi distinguer deux opérations souvent confondues. `Requery` ré-exécute entièrement la source du formulaire (coûteux, et la position courante est perdue) ; `Refresh` se contente de relire les données du jeu courant (plus léger, position conservée). On réserve `Requery` aux cas où la composition du jeu a réellement changé — par exemple pour rafraîchir une liste déroulante après l'ajout d'une valeur (`Me.cboClient.Requery`) — et l'on évite de le déclencher inutilement, notamment dans des événements fréquents comme `Current`.

## Garder un formulaire complexe en mémoire

Rouvrir à répétition un formulaire lourd est coûteux. Pour un écran fréquemment sollicité, il peut être avantageux de le charger une fois puis de le **masquer** plutôt que de le fermer, et de simplement basculer sa visibilité. Le gain de réactivité se paie en mémoire occupée : c'est un arbitrage à évaluer.

```vba
DoCmd.OpenForm "frmComplexe", , , , , acHidden   ' chargé, masqué
' plus tard :
Forms!frmComplexe.Visible = True
```

## Manipuler les enregistrements sans réinterroger

Pour parcourir ou modifier par code les enregistrements affichés par un formulaire, on utilise son `RecordsetClone` plutôt que d'ouvrir un nouveau Recordset : on évite une interrogation supplémentaire et l'on reste synchronisé avec le formulaire.

> ℹ️ Le `RecordsetClone` est détaillé au [chapitre 9.12](../09-dao-data-access-objects/12-recordsetclone.md).

## Points clés à retenir

- Un formulaire ne va pas plus vite que son `RecordSource` : restreindre colonnes et lignes, préférer une requête sauvegardée, s'appuyer sur des index.
- Ouvrir filtré via `WhereCondition` ou différer le chargement (ouvrir vide, charger après saisie des critères) évite de rapatrier des données inutiles.
- `DataEntry = True` ouvre instantanément un formulaire de saisie pure, sans charger l'existant.
- Limiter la source des listes déroulantes et les peupler à la demande ; différer le chargement des sous-formulaires (affectation du `SourceObject` à l'accès à l'onglet).
- Bannir les fonctions de domaine dans les contrôles d'un formulaire continu : déplacer le calcul dans la source via jointure ou agrégation.
- Suspendre le repeint (`Painting`) pendant les manipulations massives ; distinguer `Requery` (lourd) de `Refresh` (léger) et éviter les rafraîchissements superflus.
- Pour un écran lourd très utilisé, envisager de le masquer plutôt que de le fermer ; manipuler les enregistrements via `RecordsetClone`.

---


⏭️ [18.6. Indexation et impact sur les performances SQL](/18-optimisation-performance/06-indexation-performances-sql.md)
