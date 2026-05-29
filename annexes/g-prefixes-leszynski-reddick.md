🔝 Retour au [Sommaire](/SOMMAIRE.md)

# G. Préfixes de nommage Leszynski/Reddick pour Access

Cette annexe liste les préfixes (tags) de la convention de nommage **Leszynski/Reddick**, la plus répandue dans le monde Access. Elle complète le **chapitre 24.1**, qui en explique l'intérêt et l'application ; elle en constitue la liste de référence.

Ces conventions sont issues de deux travaux complémentaires : la **Reddick VBA Naming Convention** (RVBA, par Greg Reddick), générale pour VBA, et la **Leszynski Naming Convention** (LNC, par Stan Leszynski), spécialisée pour Access. Leur objectif : identifier d'un coup d'œil **le type et la portée** d'un objet ou d'une variable, dans les nombreux contextes où le nom apparaît seul (volet de navigation, expressions, SQL, code).

> ℹ️ Les tags exacts varient légèrement selon les sources. **La cohérence au sein d'un projet prime sur le choix précis** de tel ou tel tag.

## Structure d'un nom

```
[préfixe(s)] tag NomDeBase [Qualificatif]
```

| Composant | Rôle | Casse |
|---|---|---|
| **Préfixe(s)** | Portée et durée de vie (variables) | minuscule |
| **Tag** | Type d'objet ou de donnée (2 à 4 lettres) | minuscule |
| **NomDeBase** | Description signifiante | PascalCase |
| **Qualificatif** | Précision (`Min`, `Max`, `New`, `Old`…) | PascalCase |

Exemples :

- `tblClients` → table « Clients »
- `txtNomFamille` → zone de texte « NomFamille »
- `mstrUtilisateur` → variable de **m**odule, **str**ing, « Utilisateur »
- `gcurTauxTVA` → variable **g**lobale, **cur**rency, « TauxTVA »
- `rstCommandes` → **r**ecord**s**et DAO « Commandes »

---

## Préfixes de portée et de durée de vie

À placer **devant le tag**, pour les variables :

| Préfixe | Signification | Exemple |
|---|---|---|
| *(aucun)* | Variable locale (procédure) | `strNom` |
| `m` | Niveau module (privée au module) | `mlngCompteur` |
| `g` | Globale / publique (toute l'application) | `gblnModeAdmin` |
| `s` | Statique (conserve sa valeur entre les appels) | `sintAppels` |
| `a` | Tableau (array) | `astrNoms` |

> La RVBA prévoit des préfixes complémentaires (paramètre, passage par référence/valeur…), rarement utilisés en pratique.

---

## Tags des objets de base de données

| Tag | Objet |
|---|---|
| `tbl` | Table |
| `qry` | Requête (générique) |
| `frm` | Formulaire |
| `fsub` | Sous-formulaire |
| `rpt` | État |
| `rsub` | Sous-état |
| `mcr` | Macro |
| `bas` / `mod` | Module standard |
| `cls` | Module de classe |

### Sous-types de requêtes

| Tag | Type de requête |
|---|---|
| `qsel` | Sélection |
| `qapp` | Ajout (append) |
| `qdel` | Suppression |
| `qmak` | Création de table (make-table) |
| `qupd` | Mise à jour |
| `qxtb` | Analyse croisée (crosstab) |
| `quni` | Union |
| `qspt` | Pass-Through |
| `qddl` | Définition de données (DDL) |

---

## Tags des contrôles (formulaires et états)

| Tag | Contrôle |
|---|---|
| `lbl` | Étiquette (Label) |
| `txt` | Zone de texte (TextBox) |
| `cbo` | Liste déroulante (ComboBox) |
| `lst` | Liste (ListBox) |
| `cmd` | Bouton de commande (CommandButton) |
| `chk` | Case à cocher (CheckBox) |
| `opt` | Bouton d'option (OptionButton) |
| `ogp` | Groupe d'options (OptionGroup) |
| `tgl` | Bouton bascule (ToggleButton) |
| `tab` | Contrôle onglet (TabControl) |
| `pge` | Page d'onglet |
| `sub` | Contrôle sous-formulaire / sous-état |
| `img` | Image |
| `ole` | Cadre d'objet OLE (lié ou indépendant) |
| `ocx` | Contrôle ActiveX |
| `lin` | Trait (Line) |
| `box` | Rectangle |
| `brk` | Saut de page |

> Variantes courantes : `grp`/`fra` pour le groupe d'options, `btn` pour un bouton, `shp` pour un rectangle.

---

## Tags des variables (par type de données)

| Tag | Type VBA |
|---|---|
| `bln` | Boolean |
| `byt` | Byte |
| `int` | Integer |
| `lng` | Long |
| `cur` | Currency |
| `sng` | Single |
| `dbl` | Double |
| `dec` | Decimal |
| `dat` | Date/Time (variante `dtm`) |
| `str` | String |
| `var` | Variant (variante `vnt`) |
| `obj` | Object (générique) |

---

## Tags des objets DAO et ADO

### DAO

| Tag | Objet |
|---|---|
| `dbs` | Database |
| `wsp` | Workspace |
| `rst` | Recordset |
| `tdf` | TableDef |
| `qdf` | QueryDef |
| `fld` | Field |
| `idx` | Index |
| `rel` | Relation |
| `prp` | Property |
| `doc` | Document |
| `cnt` | Container |
| `err` | Error |

### ADO

| Tag | Objet |
|---|---|
| `cnn` | Connection (variantes `cn`, `con`) |
| `rst` | Recordset |
| `cmd` | Command |
| `prm` | Parameter |
| `fld` | Field |

---

## Constantes et énumérations

Deux usages coexistent :

- Constantes nommées en **MAJUSCULES_AVEC_SOULIGNEMENTS** (style hérité du C, fréquent pour les constantes globales) : `const TVA_STANDARD As Double = 0.2`.
- Ou avec le tag LNC `con` : `conTvaStandard`.

Comme pour le reste, l'essentiel est de **s'en tenir à un seul style** dans le projet.

---

## Exemple récapitulatif

```vba
' Variable globale (g) String (str)
Public gstrNomSociete As String

Private Sub cmdArchiver_Click()
    ' Variable locale Long (lng) et Recordset DAO (rst)
    Dim lngNbLignes As Long
    Dim rstCommandes As DAO.Recordset

    Set rstCommandes = CurrentDb.OpenRecordset("qselCommandesAnciennes")
    ' … traitement …
    rstCommandes.Close
    Set rstCommandes = Nothing
End Sub
```

Noms d'objets associés : `tblCommandes`, `qappArchiveCommandes`, `frmCommandeSaisie`, `txtMontant`, `cboPays`, `rptFactureMensuelle`.

---

## Bonnes pratiques et perspective actuelle

- **Tag en minuscules, nom de base en PascalCase**, sans espace ni caractère accentué — un espace impose des crochets `[ ]` partout (volet, SQL, expressions).
- **Champs de table** : le plus souvent nommés simplement (`Nom`, `DateInscription`), sans tag de type, contrairement aux contrôles et variables.
- **Éviter les mots réservés** comme noms (`Name`, `Date`, `Value`…), qui provoquent des ambiguïtés (chapitre 13.1).

**Le débat sur l'utilité des tags est nuancé :**

- Les tags d'**objets de base de données** (`tbl`, `qry`, `frm`, `rpt`) et de **contrôles** (`txt`, `cbo`, `cmd`) restent **largement recommandés en Access** : ces noms apparaissent isolément dans le volet de navigation, le SQL et les expressions, où le type ne se déduit pas autrement.
- Les tags de **type de variable** (`str`, `lng`…) sont plus discutés : avec le typage fort et l'auto-complétion de l'éditeur, certains les jugent redondants. D'autres y voient un repère utile à la relecture. Les deux positions se défendent ; le critère reste la cohérence.

---

> 💡 **À retenir.** La convention se résume à une grammaire simple : un éventuel **préfixe de portée**, un **tag de type** en minuscules, puis un **nom de base** en PascalCase. En Access, les tags d'objets et de contrôles apportent une valeur réelle de lisibilité. Pour l'application concrète de ces règles à l'architecture d'un projet, voir les chapitres 24.1 et 24.2.

⏭️ [H. Déclarations API Windows compatibles 32/64 bits](/annexes/h-declarations-api-32-64-bits.md)
