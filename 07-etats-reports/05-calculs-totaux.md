🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 7.5. Calculs et totaux dans les états

Les états sont l'outil privilégié pour **synthétiser** des données : sous-totaux par groupe, totaux généraux, moyennes, comptages, cumuls progressifs. Access offre plusieurs mécanismes pour les calculer, du plus simple (une expression dans un contrôle) au plus souple (l'accumulation en VBA). Cette section les passe en revue, en indiquant lequel choisir selon le besoin, et livre enfin le pilotage des regroupements et du tri par la collection `GroupLevel`, annoncé aux sections 7.3 et 7.4.

---

## 7.5.1. Les approches du calcul dans un état

| Approche | Principe | Quand l'employer |
|---|---|---|
| Expression dans la source d'un contrôle | `=expression` | Calcul ligne à ligne ou simple combinaison |
| Fonctions d'agrégat (`Somme`, `Moyenne`…) | Placées dans la bonne section | Sous-totaux et totaux généraux |
| Propriété `RunningSum` | Cumul d'un contrôle au fil des lignes | Totaux courants, numérotation |
| Accumulation en VBA | Variable cumulée dans les événements | Logique non exprimable en agrégat |
| Fonctions de domaine (`DSum`, `DCount`) | Agrégat depuis une **autre** source | Valeur extérieure à la source de l'état |

Le bon réflexe : commencer par les expressions et les agrégats, qui couvrent l'essentiel sans code.

---

## 7.5.2. Contrôles calculés par expression

Un contrôle dont la source (`ControlSource`) commence par `=` affiche le résultat d'une **expression** plutôt qu'un champ. Dans la section Détail, l'expression se calcule pour chaque enregistrement :

```
' Source d'un contrôle de la section Détail
=[Quantite]*[PrixUnitaire]
```

> **Localisation** : dans les expressions de contrôle, les noms de fonctions et le séparateur d'arguments suivent les paramètres régionaux (par exemple `Somme`, `Nz([Champ];0)` en configuration française). En VBA, ce sont les noms anglais et la virgule. Ce point est développé au chapitre 11.7.

---

## 7.5.3. Fonctions d'agrégat et placement par section

Les fonctions d'agrégat — `Somme`, `Moyenne`, `Compte`, `Min`, `Max`, `EcType`, `Var`, `Premier`, `Dernier` — synthétisent un ensemble de lignes. **Leur portée dépend de la section** où on les place (rappel et approfondissement de 7.3) :

```
' Sous-total d'un groupe → dans un pied de groupe
=Somme([MontantLigne])

' Total général → dans un pied d'état
=Somme([MontantLigne])

' Comptage
=Compte(*)            ' nombre d'enregistrements
=Compte([Email])      ' nombre de valeurs non nulles
```

| Section | Portée de l'agrégat |
|---|---|
| En-tête / pied de groupe | Le **groupe** courant |
| En-tête / pied d'état | **Tous** les enregistrements |
| En-tête / pied de **page** | **Interdit** → `#Erreur` |

C'est l'**emplacement seul** qui détermine si l'on obtient un sous-total ou un total général ; aucune option supplémentaire n'est requise.

---

## 7.5.4. La propriété RunningSum : cumuls et numérotation

Les fonctions d'agrégat donnent un total *global* par section, mais pas un **cumul progressif** ligne après ligne. C'est le rôle de la propriété **`RunningSum`** d'un contrôle, propre aux états. Elle prend trois valeurs :

- **Aucun** : pas de cumul (comportement normal) ;
- **Sur le groupe** : cumul réinitialisé à chaque nouveau groupe ;
- **Sur tout** : cumul sur l'ensemble de l'état.

Appliquée à un contrôle lié à une valeur numérique, elle accumule cette valeur au fil de la section Détail :

```
' Contrôle txtCumul : Source = =[Montant], RunningSum = Sur le groupe
' → affiche le cumul des montants, remis à zéro à chaque groupe
```

`RunningSum` est la **manière idiomatique** d'obtenir un total courant dans un état ; il serait malaisé de le reproduire avec des fonctions d'agrégat.

---

## 7.5.5. Numérotation des lignes

Cas particulier élégant de `RunningSum` : pour numéroter les lignes, on lie un contrôle à la constante `=1` et on active le cumul. Le contrôle additionne alors `1` à chaque ligne, produisant `1, 2, 3…` :

```
' Contrôle txtNumLigne : Source = =1, RunningSum = Sur tout
' → numérotation continue : 1, 2, 3, ...

' Avec RunningSum = Sur le groupe
' → numérotation repartant de 1 à chaque groupe
```

C'est la technique standard, bien plus fiable qu'un compteur géré à la main.

---

## 7.5.6. Totaux conditionnels : privilégier la requête

Pour un total qui ne porte que sur les lignes remplissant une **condition** (compter les articles en rupture, sommer les montants impayés…), la solution la plus robuste **n'est pas** le VBA, mais la **requête source**. On y ajoute une colonne calculée qui vaut 1/0 (ou le montant/0) selon la condition, puis on la somme dans la section voulue :

```sql
-- Colonne ajoutée à la requête source de l'état
SELECT *, IIf([Stock] < [Seuil], 1, 0) AS EnAlerte
FROM T_Articles
```

```
' Dans le pied de groupe ou d'état
=Somme([EnAlerte])      ' nombre d'articles en alerte
```

Cette approche contourne tous les pièges de calendrier des événements `Format`/`Print` (section 7.2) : le calcul est fait par le moteur, indépendamment de la mise en page.

---

## 7.5.7. Accumulation en VBA dans les événements

Lorsqu'une logique ne peut s'exprimer ni en agrégat ni en colonne de requête, on accumule dans une **variable de module**, mise à jour au fil des événements. Deux précautions découlent du fonctionnement vu au chapitre 7.2 :

- **réinitialiser** l'accumulateur dans l'en-tête de groupe (pour un cumul par groupe) ;
- **accumuler dans `Print`** plutôt que dans `Format`, afin de ne compter que ce qui est réellement imprimé, et **compenser dans `Retreat`** si Access revient en arrière.

```vba
' En tête du module de l'état
Private m_curCumul As Currency

Private Sub GroupHeader0_Format(Cancel As Integer, FormatCount As Integer)
    If FormatCount = 1 Then m_curCumul = 0   ' réinitialiser au début du groupe
End Sub

Private Sub Detail_Print(Cancel As Integer, PrintCount As Integer)
    ' Cumuler uniquement les lignes effectivement imprimées
    m_curCumul = m_curCumul + Nz(Me.txtMontant.Value, 0)
End Sub
```

> Cette voie reste un **dernier recours** : pour la plupart des besoins, l'expression, l'agrégat, `RunningSum` ou la colonne de requête sont plus simples et plus sûrs.

---

## 7.5.8. Fonctions de domaine (DSum, DCount) dans un état

Lorsqu'un calcul porte sur une **autre** source que celle de l'état (une table non jointe, par exemple un cumul historique), on recourt aux **fonctions de domaine** : `DSum`, `DCount`, `DLookup`, etc. (chapitre 11.10).

```
' Source d'un contrôle : total des règlements d'un client, depuis une autre table
=DSomme("Montant";"T_Reglements";"ClientID=" & [ClientID])
```

Ces fonctions sont **coûteuses** : placées dans la section Détail, elles s'exécutent une fois **par enregistrement** et peuvent fortement ralentir l'état. À réserver aux en-têtes/pieds ou aux cas où la jointure dans la requête source est impossible.

---

## 7.5.9. Regroupement et tri par code : la collection GroupLevel

Les regroupements et le tri d'un état sont représentés par la collection **`GroupLevel`** (annoncée en 7.3 et 7.4). Chaque niveau expose notamment :

- **`ControlSource`** : le champ ou l'expression sur lequel on regroupe/trie ;
- **`SortOrder`** : l'ordre — `False` (croissant) ou `True` (décroissant) ;
- **`GroupOn`** / **`GroupInterval`** : le mode de regroupement (chaque valeur, par préfixe, par intervalle…) ;
- **`KeepTogether`** : la cohésion du groupe face aux sauts de page.

L'indexation est **basée sur 0** : `GroupLevel(0)` est le groupe le plus externe (rappel du piège de numérotation, section 7.3).

Pour **trier dynamiquement** un état, on modifie ces propriétés dans l'événement `Open` — par exemple selon un choix de l'utilisateur :

```vba
Private Sub Report_Open(Cancel As Integer)
    ' Le niveau GroupLevel(0) doit déjà exister (défini en mode création)
    Me.GroupLevel(0).ControlSource = Nz(Forms!F_Criteres!cboTri.Value, "Nom")
    Me.GroupLevel(0).SortOrder = False   ' croissant
End Sub
```

> **Limite** : on modifie ici un niveau **existant**. Pour **créer** des niveaux de regroupement par code, on utilise la fonction `CreateGroupLevel`, abordée avec la génération d'états programmatique au chapitre 7.7.

C'est la collection `GroupLevel` (ou la clause `ORDER BY` du `RecordSource`) qui pilote réellement le tri d'un état — et non la propriété `OrderBy`, souvent ignorée (chapitre 7.4).

---

## 7.5.10. Pièges et bonnes pratiques

- **Placer les agrégats dans la bonne section** : groupe ou état, jamais page (`#Erreur`).
- **Protéger contre `Null`** avec `Nz` dans les calculs : un champ vide propage `#Erreur` ou fausse une somme.
- **Se prémunir de la division par zéro** dans les expressions (tester le dénominateur avant de diviser).
- **`Compte(*)` vs `Compte([champ])`** : le second ignore les valeurs nulles.
- **Préférer `RunningSum`** aux compteurs manuels pour les cumuls et la numérotation.
- **Pour un total conditionnel, passer par la requête** (colonne `IIf`) plutôt que par le VBA.
- **En VBA, accumuler dans `Print`**, réinitialiser dans l'en-tête de groupe, compenser dans `Retreat` (chapitre 7.2).
- **Limiter les fonctions de domaine** dans la section Détail : elles s'exécutent par enregistrement et pénalisent les performances.
- **Trier via `GroupLevel` ou la requête**, pas via `OrderBy` (chapitre 7.4).

---

## 7.5.11. Récapitulatif

- Les calculs d'un état reposent d'abord sur des **expressions** (`=…`) et des **fonctions d'agrégat**, dont la **portée dépend de la section** (groupe, état ; jamais page).
- La propriété **`RunningSum`** fournit les **cumuls progressifs** et la **numérotation des lignes** (contrôle lié à `=1`), de façon bien plus fiable qu'un compteur manuel.
- Pour un **total conditionnel**, on ajoute une colonne calculée à la **requête source** puis on la somme, ce qui évite les pièges de calendrier des événements.
- L'**accumulation en VBA** est un dernier recours : accumuler dans `Print`, réinitialiser dans l'en-tête de groupe, compenser dans `Retreat` (chapitre 7.2).
- Les **fonctions de domaine** (`DSum`, `DCount`… — chapitre 11.10) agrègent depuis une autre source, mais sont coûteuses dans la section Détail.
- La collection **`GroupLevel`** (indexée à partir de 0) pilote regroupement et tri ; on modifie un niveau existant dans `Open`, et l'on crée des niveaux par `CreateGroupLevel` (chapitre 7.7). C'est elle — ou `ORDER BY` — qui gouverne le tri, pas `OrderBy`.

⏭️ [7.6. Sous-états et liaison parent/enfant](/07-etats-reports/06-sous-etats.md)
