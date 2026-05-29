🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 6.6. Propriété OpenArgs — passage de paramètres à l'ouverture

Lorsqu'un formulaire en ouvre un autre, il a souvent besoin de lui transmettre un contexte : un mode de fonctionnement (« ajout » ou « modification »), l'identifiant d'un enregistrement à afficher, ou toute autre information utile au démarrage. La propriété **`OpenArgs`** est le mécanisme natif d'Access conçu pour cela : elle permet de passer une **chaîne de caractères** au formulaire (ou à l'état) au moment de son ouverture.

Cette section décrit comment définir et lire `OpenArgs`, comment transmettre plusieurs valeurs, les cas d'usage classiques, ainsi que les alternatives lorsque cette propriété atteint ses limites.

---

## 6.6.1. Principe et place dans DoCmd.OpenForm

`OpenArgs` est le **dernier argument** de la méthode `DoCmd.OpenForm`. La signature complète est la suivante :

```vba
DoCmd.OpenForm FormName, View, FilterName, WhereCondition, _
               DataMode, WindowMode, OpenArgs
```

| Argument | Rôle |
|---|---|
| `FormName` | Nom du formulaire à ouvrir (obligatoire) |
| `View` | Mode d'affichage (`acNormal`, `acDesign`…) |
| `FilterName` | Requête servant de filtre |
| `WhereCondition` | Clause `WHERE` (sans le mot-clé) |
| `DataMode` | Mode de saisie (`acFormAdd`, `acFormEdit`, `acFormReadOnly`) |
| `WindowMode` | Mode de fenêtre (`acWindowNormal`, `acDialog`, `acHidden`…) |
| `OpenArgs` | **Chaîne transmise au formulaire ouvert** |

Comme `OpenArgs` se situe en dernière position, on utilise des **virgules de remplacement** pour sauter les arguments intermédiaires non utilisés :

```vba
' Ouvrir un formulaire en ne renseignant que OpenArgs
DoCmd.OpenForm "F_Detail", , , , , , "42"
```

---

## 6.6.2. Définir OpenArgs à l'ouverture

La valeur transmise est toujours interprétée comme une chaîne. Voici plusieurs exemples typiques depuis le formulaire appelant :

```vba
' Transmettre un mode de fonctionnement
DoCmd.OpenForm "F_Client", , , , acFormAdd, , "NouveauClient"

' Transmettre l'identifiant d'un enregistrement
DoCmd.OpenForm "F_DetailCommande", , , , , , Me.txtCommandeID.Value

' Transmettre une valeur en combinant avec un filtre
DoCmd.OpenForm "F_Factures", , , "ClientID = " & Me.ClientID, , , "DepuisFicheClient"
```

> **Remarque** : même lorsqu'on transmet un nombre (un identifiant, par exemple), Access le reçoit sous forme de texte. Cela ne pose pas de problème pour une concaténation SQL, mais nécessite une conversion explicite (`CLng`, `CInt`) si l'on veut l'utiliser dans un calcul arithmétique.

---

## 6.6.3. Lire OpenArgs dans le formulaire ouvert (et gérer Null)

Dans le formulaire ouvert, la valeur se lit via `Me.OpenArgs`. **Point essentiel** : si aucun argument n'a été transmis, `Me.OpenArgs` renvoie **`Null`**. Toute opération de chaîne sur une valeur `Null` échoue ou propage `Null` ; il faut donc systématiquement tester ce cas.

```vba
Private Sub Form_Open(Cancel As Integer)
    If Not IsNull(Me.OpenArgs) Then
        Select Case Me.OpenArgs
            Case "NouveauClient"
                Me.Caption = "Création d'un client"
            Case "Modification"
                Me.Caption = "Modification d'un client"
            Case Else
                Me.Caption = "Fiche client"
        End Select
    End If
End Sub
```

La fonction `Nz` offre une alternative concise pour fournir une valeur par défaut :

```vba
Dim strMode As String
strMode = Nz(Me.OpenArgs, "Consultation")   ' "Consultation" si OpenArgs est Null
```

---

## 6.6.4. Passer plusieurs valeurs : délimiteurs et Split

`OpenArgs` ne transporte qu'**une seule chaîne**. Pour transmettre plusieurs informations, on les concatène avec un **séparateur** convenu, puis on les sépare à la réception avec la fonction `Split`.

```vba
' --- Formulaire appelant : un mode et un identifiant séparés par ";" ---
DoCmd.OpenForm "F_Detail", , , , , , "Modification;42"
```

```vba
' --- Formulaire ouvert : analyse de la chaîne ---
Private Sub Form_Open(Cancel As Integer)
    Dim arrArgs() As String

    If Not IsNull(Me.OpenArgs) Then
        arrArgs = Split(Me.OpenArgs, ";")
        ' arrArgs(0) = "Modification"
        ' arrArgs(1) = "42"

        Me.Caption = arrArgs(0)
    End If
End Sub
```

Pour un nombre d'arguments variable ou nommés, un format **clé=valeur** est plus robuste et auto-documenté :

```vba
' Appel : "mode=edit;id=42;source=client"
DoCmd.OpenForm "F_Detail", , , , , , "mode=edit;id=42;source=client"
```

```vba
' Réception : conversion en Dictionary pour un accès par clé
Private Function ParseOpenArgs(ByVal s As String) As Object
    Dim dic As Object, arrPaires() As String, arrKV() As String
    Dim i As Long

    Set dic = CreateObject("Scripting.Dictionary")
    If Len(s) > 0 Then
        arrPaires = Split(s, ";")
        For i = LBound(arrPaires) To UBound(arrPaires)
            arrKV = Split(arrPaires(i), "=")
            If UBound(arrKV) = 1 Then dic(arrKV(0)) = arrKV(1)
        Next i
    End If
    Set ParseOpenArgs = dic
End Function
```

L'objet `Scripting.Dictionary` est détaillé au chapitre 3.7.

---

## 6.6.5. Cas d'usage courants

### Définir un mode de fonctionnement

Le cas le plus fréquent : indiquer au formulaire s'il s'ouvre en création, en modification ou en consultation, afin d'adapter son titre, ses contrôles ou ses boutons.

### Positionner le formulaire sur un enregistrement

`OpenArgs` transmet l'identifiant, et le formulaire se place sur la ligne correspondante grâce au `RecordsetClone` et à la propriété `Bookmark`. Cette opération nécessite que le jeu d'enregistrements soit chargé : elle se réalise donc dans l'événement **`Load`**.

```vba
Private Sub Form_Load()
    Dim rs As DAO.Recordset

    If Not IsNull(Me.OpenArgs) Then
        Set rs = Me.RecordsetClone
        rs.FindFirst "ClientID = " & Me.OpenArgs
        If Not rs.NoMatch Then
            Me.Bookmark = rs.Bookmark
        End If
        Set rs = Nothing
    End If
End Sub
```

Le `RecordsetClone` et la propriété `Bookmark` sont étudiés au chapitre 9.12.

### Pré-remplir un contrôle

À l'ouverture d'un formulaire de saisie, on peut pré-renseigner un champ avec la valeur transmise (par exemple la ville par défaut héritée du formulaire appelant) :

```vba
Private Sub Form_Load()
    If Not IsNull(Me.OpenArgs) Then
        Me.txtVille.Value = Me.OpenArgs
    End If
End Sub
```

---

## 6.6.6. OpenArgs et les états (Reports)

Les états disposent eux aussi de la propriété `OpenArgs`. Avec `DoCmd.OpenReport`, l'argument occupe la **sixième** position (les états n'ont pas d'argument `DataMode`) :

```vba
DoCmd.OpenReport ReportName, View, FilterName, WhereCondition, _
                 WindowMode, OpenArgs
```

```vba
' Ouvrir un état en lui transmettant un libellé de période
DoCmd.OpenReport "E_Ventes", acViewPreview, , , , "Trimestre 1"
```

```vba
' Réception dans l'état
Private Sub Report_Open(Cancel As Integer)
    If Not IsNull(Me.OpenArgs) Then
        Me.lblPeriode.Caption = Me.OpenArgs
    End If
End Sub
```

Les événements et les sections des états sont traités au chapitre 7.

---

## 6.6.7. Quand lire OpenArgs : Open ou Load ?

`OpenArgs` est renseignée **avant** le déclenchement du premier événement du formulaire. Elle est donc disponible dès l'événement `Open`, puis tout au long de la vie du formulaire.

| Action souhaitée | Événement recommandé |
|---|---|
| Adapter le titre, annuler l'ouverture (`Cancel`) | `Form_Open` |
| Lire un simple drapeau de mode | `Form_Open` |
| Pré-remplir un contrôle, positionner sur un enregistrement | `Form_Load` (jeu et contrôles disponibles) |

Le cycle complet des événements d'un formulaire (ordre `Open` → `Load` → …) est détaillé au chapitre 8.7.

---

## 6.6.8. Alternatives à OpenArgs

`OpenArgs` est simple et fiable, mais limitée à une chaîne. D'autres mécanismes la complètent :

- **`TempVars`** — variables globales de session, accessibles depuis n'importe quel objet sans couplage entre formulaires. Utiles pour un contexte partagé par plusieurs écrans (utilisateur courant, paramètres applicatifs). Voir chapitre 15.7.
- **Propriétés publiques du module de classe du formulaire** — on déclare des `Property Let/Get` dans le formulaire ouvert, puis on les affecte depuis l'appelant après ouverture. Cette approche permet de transmettre des types non textuels (objets, collections) et de communiquer dans les deux sens. Voir chapitres 6.9 et 16.2.

**Règle pratique** : `OpenArgs` convient pour un à quelques paramètres simples au démarrage. Au-delà, ou pour transmettre autre chose qu'une chaîne, privilégier les `TempVars` ou les propriétés de classe.

---

## 6.6.9. Pièges et bonnes pratiques

- **Toujours tester `IsNull(Me.OpenArgs)`** (ou utiliser `Nz`) avant toute opération de chaîne : l'absence d'argument renvoie `Null`, pas une chaîne vide.
- **La valeur est toujours du texte** : convertir explicitement (`CLng`, `CDate`…) avant tout usage numérique ou de date.
- **Choisir un séparateur absent des données** lorsqu'on encode plusieurs valeurs (préférer un caractère peu courant comme `|` ou `;` si les données n'en contiennent pas).
- **Documenter le format attendu** par un commentaire dans l'en-tête du module du formulaire ouvert (par exemple : *« OpenArgs attendu : mode=edit;id=NNN »*).
- **Ne pas surcharger `OpenArgs`** : pour de nombreux paramètres ou des objets, basculer vers `TempVars` ou des propriétés de classe.
- **Positionner sur un enregistrement dans `Load`**, jamais dans `Open` : le jeu d'enregistrements n'est pas encore disponible à l'ouverture.

---

## 6.6.10. Récapitulatif

- `OpenArgs` est le dernier argument de `DoCmd.OpenForm` (sixième pour `DoCmd.OpenReport`) ; il transmet une **chaîne** au formulaire ou à l'état ouvert.
- On utilise des **virgules de remplacement** pour ne renseigner que cet argument.
- La lecture se fait via `Me.OpenArgs`, qui renvoie **`Null`** en l'absence d'argument : tester `IsNull` ou employer `Nz` est impératif.
- Pour transmettre plusieurs valeurs, on les **concatène avec un séparateur** et on les analyse via `Split`, éventuellement sous forme **clé=valeur**.
- Cas d'usage typiques : mode de fonctionnement, positionnement sur un enregistrement (dans `Load`, via `RecordsetClone` et `Bookmark`), pré-remplissage de contrôles.
- Lorsque `OpenArgs` ne suffit pas (types non textuels, nombreux paramètres, communication bidirectionnelle), recourir aux **`TempVars`** (chapitre 15.7) ou aux **propriétés du module de classe** (chapitres 6.9 et 16.2).

⏭️ [6.7. Formulaires modaux et boîtes de dialogue personnalisées](/06-formulaires/07-formulaires-modaux-dialogues.md)
