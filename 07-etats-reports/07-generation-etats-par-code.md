🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 7.7. Génération d'états par code (création, modification de contrôles)

Access permet de **créer et modifier des états par programmation** : générer un état de toutes pièces, y ajouter des contrôles, définir des niveaux de regroupement. Cette capacité sert principalement à produire des états à partir de modèles ou de métadonnées, ou à fabriquer en série des états similaires. Elle obéit toutefois à une **contrainte stricte** qu'il faut comprendre avant tout : les modifications de structure ne sont possibles qu'en **mode Création**.

Cette section présente les fonctions de génération (`CreateReport`, `CreateReportControl`, `CreateGroupLevel`), la modification d'un état existant, et les limites pratiques — dont l'impossibilité d'opérer dans une base compilée en ACCDE.

---

## 7.7.1. La contrainte fondamentale : le mode Création

Toute modification **structurelle** d'un état — le créer, ajouter ou supprimer des contrôles, lier un contrôle à un champ, ajouter un niveau de regroupement — n'est possible que lorsque l'état est ouvert en **mode Création** (`acViewDesign`). On **ne peut pas** ajouter un contrôle à un état affiché en aperçu.

À distinguer des ajustements **cosmétiques** (couleur, visibilité, libellé), qui se règlent à l'exécution dans les événements `Open`/`Format` (voir 7.7.9).

> **Avertissement majeur** : dans une base compilée en **ACCDE** (chapitre 20.3), la modification de la conception des objets est **impossible**. Le code de génération d'états présenté ici **ne fonctionne donc pas** en ACCDE. Il est à réserver aux phases de développement ou aux applications distribuées en `.accdb`.

---

## 7.7.2. Créer un état : CreateReport

La fonction `CreateReport` crée un nouvel état, **ouvert en mode Création**, et en renvoie l'objet `Report` :

```vba
Dim rpt As Report
Set rpt = CreateReport()        ' nouvel état, ouvert en conception
Debug.Print rpt.Name            ' nom temporaire, du type "Report1"
```

L'état porte un **nom temporaire** (« Report1 ») jusqu'à son enregistrement (section 7.7.6). On peut éventuellement fournir un modèle d'état comme base.

---

## 7.7.3. Ajouter des contrôles : CreateReportControl

`CreateReportControl` ajoute un contrôle à un état ouvert en conception et en renvoie l'objet `Control` :

```vba
CreateReportControl(NomEtat, TypeContrôle, [Section], [Parent], _
                    [ChampLié], [Gauche], [Haut], [Largeur], [Hauteur])
```

| Argument | Rôle |
|---|---|
| `NomEtat` | Nom (chaîne) de l'état en conception |
| `TypeContrôle` | Type : `acLabel`, `acTextBox`, `acLine`, `acImage`, `acSubform`… |
| `Section` | Section cible : `acDetail`, `acPageHeader`, `acHeader`… (constantes du chapitre 7.1) |
| `Parent` | Contrôle parent (groupe d'options, onglet) ; `""` au premier niveau |
| `ChampLié` | Champ source pour un contrôle lié ; `""` si indépendant |
| `Gauche`, `Haut`, `Largeur`, `Hauteur` | Position et taille, **en twips** |

```vba
Dim ctl As Control

' Étiquette de titre dans l'en-tête de page
Set ctl = CreateReportControl(rpt.Name, acLabel, acPageHeader, , , 200, 200, 4000, 400)
ctl.Caption = "Liste des clients"

' Zone de texte liée au champ Nom, dans la section Détail
Set ctl = CreateReportControl(rpt.Name, acTextBox, acDetail, , "Nom", 200, 100, 3000, 300)
ctl.Name = "txtNom"
```

> La **section cible doit exister** dans l'état. Les sections Détail, en-tête/pied de page sont présentes par défaut ; recourir à l'en-tête/pied d'état suppose que ces sections soient activées.

---

## 7.7.4. Définir les propriétés des contrôles et de l'état

Une fois le contrôle créé, on règle ses propriétés sur l'objet renvoyé ; de même pour l'état :

```vba
' Propriétés de l'état
rpt.RecordSource = "T_Clients"
rpt.Caption = "Clients"

' Propriétés d'un contrôle
ctl.FontName = "Calibri"
ctl.FontSize = 11
ctl.TextAlign = 1            ' alignement
```

La définition de la source (`RecordSource`) suit les principes du chapitre 7.4.

---

## 7.7.5. Créer des niveaux de regroupement : CreateGroupLevel

Annoncée au chapitre 7.5, la fonction `CreateGroupLevel` ajoute un niveau de regroupement/tri à un état **en conception**. Elle renvoie le numéro du niveau créé :

```vba
CreateGroupLevel(NomEtat, Expression, EnTête, Pied)
```

- `Expression` : le champ ou l'expression de regroupement ;
- `EnTête` / `Pied` : `True` pour créer la section d'en-tête / de pied de groupe correspondante.

```vba
' Regrouper par Ville, avec en-tête de groupe mais sans pied
Dim lngNiveau As Long
lngNiveau = CreateGroupLevel(rpt.Name, "Ville", True, False)
```

Une fois le niveau créé, ses propriétés (`SortOrder`, `GroupOn`…) se manipulent via la collection `GroupLevel` (chapitre 7.5).

---

## 7.7.6. Enregistrer et nommer l'état généré

L'état généré porte un nom temporaire ; il faut l'**enregistrer** puis le **renommer** pour lui donner un nom définitif. Une séquence fiable :

```vba
Dim strNomTemp As String
strNomTemp = rpt.Name              ' ex. "Report1"

DoCmd.Save acReport, strNomTemp     ' enregistre sous le nom temporaire
DoCmd.Close acReport, strNomTemp    ' ferme le mode Création
DoCmd.Rename "E_ClientsGenere", acReport, strNomTemp   ' renomme
```

L'état est ensuite utilisable normalement, par exemple en aperçu via `DoCmd.OpenReport "E_ClientsGenere", acViewPreview`.

Exemple complet, assemblant les étapes précédentes :

```vba
Sub GenererEtatClients()
    Dim rpt As Report
    Dim ctl As Control
    Dim strNomTemp As String

    Set rpt = CreateReport()
    strNomTemp = rpt.Name
    rpt.RecordSource = "T_Clients"

    Set ctl = CreateReportControl(strNomTemp, acLabel, acPageHeader, , , 200, 200, 4000, 400)
    ctl.Caption = "Liste des clients"
    ctl.FontSize = 14

    Set ctl = CreateReportControl(strNomTemp, acTextBox, acDetail, , "Nom", 200, 100, 3000, 300)
    ctl.Name = "txtNom"

    CreateGroupLevel strNomTemp, "Ville", True, False

    DoCmd.Save acReport, strNomTemp
    DoCmd.Close acReport, strNomTemp
    DoCmd.Rename "E_ClientsGenere", acReport, strNomTemp
End Sub
```

---

## 7.7.7. Modifier un état existant par code

Pour modifier un état **déjà conçu**, on l'ouvre en mode Création, on agit sur ses contrôles, puis on enregistre et on ferme :

```vba
DoCmd.OpenReport "E_Clients", acViewDesign

' L'état est maintenant accessible via la collection Reports
Reports!E_Clients.RecordSource = "T_ClientsActifs"
Reports!E_Clients.txtNom.FontSize = 12

DoCmd.Close acReport, "E_Clients", acSaveYes
```

C'est l'équivalent, pour les états, de la manipulation dynamique des contrôles de formulaire traitée au chapitre 8.11 — mais en mode Création pour les changements de structure.

---

## 7.7.8. Supprimer un contrôle : DeleteReportControl

La fonction `DeleteReportControl` retire un contrôle d'un état ouvert en conception :

```vba
DoCmd.OpenReport "E_Clients", acViewDesign
DeleteReportControl "E_Clients", "txtAncienChamp"
DoCmd.Close acReport, "E_Clients", acSaveYes
```

---

## 7.7.9. Propriétés modifiables à l'exécution vs en conception

Toutes les propriétés ne relèvent pas du mode Création :

| Type de modification | Où l'opérer |
|---|---|
| Ajouter/supprimer un contrôle, lier un contrôle (`ControlSource`), renommer, repositionner | **Mode Création** |
| Ajouter un niveau de regroupement | **Mode Création** |
| Visibilité (`Visible`), couleurs (`ForeColor`, `BackColor`), libellé d'étiquette (`Caption`), valeur affichée | **À l'exécution** (`Open`, `Format`) |

Autrement dit, la plupart des **adaptations cosmétiques** d'un état existant se font sans mode Création, dans les événements (chapitres 7.2 et 17.8) ; seule la **structure** impose la conception.

---

## 7.7.10. Quand (ne pas) générer des états par code

Générer un état entier par code est **rare, verbeux et fragile**. Avant de s'y engager, il faut vérifier qu'une approche plus simple ne suffit pas :

- **source et filtres dynamiques** sur un état conçu visuellement (chapitre 7.4) ;
- **tri dynamique** via `GroupLevel` (chapitre 7.5) ;
- **mise en forme conditionnelle** par code ou intégrée (chapitre 17.8) ;
- **sous-état changé dynamiquement** par `SourceObject` (chapitre 7.6).

La génération par code se justifie surtout pour : des **modèles d'états définis par l'utilisateur**, la **production en série** d'états très similaires, ou la **construction à partir de métadonnées**. Pour le reste, la conception visuelle assortie d'ajustements à l'exécution est préférable.

---

## 7.7.11. Pièges et bonnes pratiques

- **Mode Création obligatoire** pour toute modification de structure : on ne modifie pas la conception d'un état en aperçu.
- **Inopérant en ACCDE** : ne pas concevoir d'application reposant sur la génération d'états à l'exécution si elle est distribuée en ACCDE (chapitres 20.3 et 21.1).
- **Positions et tailles en twips** (1 cm ≈ 567 twips).
- **Vérifier l'existence de la section cible** avant d'y ajouter un contrôle.
- **Enregistrer puis renommer** : l'état généré porte un nom temporaire (« Report1 ») jusqu'à son enregistrement.
- **Fermer le mode Création** avant d'ouvrir l'état en aperçu.
- **Préférer les approches dynamiques** (chapitres 7.4, 7.5, 7.6, 17.8) à la génération complète, sauf besoin réel de templating.

---

## 7.7.12. Récapitulatif

- La modification **structurelle** d'un état (création, ajout/suppression de contrôles, regroupements) exige le **mode Création** ; elle est **impossible en ACCDE** (chapitre 20.3).
- `CreateReport` crée un état en conception ; `CreateReportControl` y ajoute des contrôles (type, section, champ lié, position en twips) ; `CreateGroupLevel` ajoute des niveaux de regroupement (chapitre 7.5).
- Un état généré porte un nom temporaire : il faut l'**enregistrer** (`DoCmd.Save`) puis le **renommer** (`DoCmd.Rename`).
- Un état existant se modifie en l'ouvrant en `acViewDesign`, en agissant sur ses contrôles via la collection `Reports`, puis en le fermant avec sauvegarde ; `DeleteReportControl` supprime un contrôle.
- Les modifications **cosmétiques** (visibilité, couleurs, libellés) se font à l'exécution dans les événements ; seule la structure impose la conception (chapitres 7.2 et 17.8).
- La génération complète par code reste un cas particulier (templating, production en série) : pour la plupart des besoins, conception visuelle + source/tri/mise en forme dynamiques (chapitres 7.4 à 7.6, 17.8) suffisent.

⏭️ [7.8. Export d'états (PDF, RTF, HTML) par DoCmd.OutputTo](/07-etats-reports/08-export-etats.md)
