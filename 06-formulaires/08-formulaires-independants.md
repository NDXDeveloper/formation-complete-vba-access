🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 6.8. Formulaires indépendants (unbound forms) pour saisies complexes

Un formulaire Access est dit **lié** (bound) lorsqu'il puise ses données dans une source via sa propriété `RecordSource`, et que ses contrôles sont rattachés à des champs par leur `ControlSource`. Access prend alors tout en charge automatiquement : lecture, affichage, navigation, enregistrement. À l'inverse, un formulaire **indépendant** (unbound) n'a pas de source : c'est le développeur qui pilote intégralement le chargement, la modification et l'enregistrement des données par code.

Cette section explique quand et pourquoi recourir à un formulaire indépendant, son cycle de vie complet (charger → modifier → enregistrer), et les patrons de code qui remplacent les automatismes perdus.

---

## 6.8.1. Formulaire lié ou indépendant : la distinction

| Aspect | Formulaire lié | Formulaire indépendant |
|---|---|---|
| `RecordSource` | Renseigné (table, requête, SQL) | Vide |
| `ControlSource` des contrôles | Rattaché à un champ | Vide |
| Lecture des données | Automatique | Par code (Recordset, SQL) |
| Enregistrement | Automatique | Par code explicite |
| Navigation | Automatique | À réimplémenter |
| Verrouillage / suivi des modifications | Géré par Access | À réimplémenter |

Dans un formulaire indépendant, les contrôles ne sont que des **conteneurs en mémoire** : leur contenu n'a aucun lien direct avec la base tant que le code ne l'établit pas.

---

## 6.8.2. Pourquoi recourir à un formulaire indépendant

Le surcroît de code se justifie dans plusieurs situations :

- **Saisie répartie sur plusieurs tables** : assembler une transaction (en-tête + lignes, par exemple) et ne l'écrire qu'une fois cohérente.
- **Validation avant écriture** : rien n'est enregistré tant que l'utilisateur n'a pas confirmé ; on évite les enregistrements partiels ou incohérents en base.
- **Assistants multi-étapes** : collecter des informations sur plusieurs écrans avant de tout valider en bloc.
- **Sources de données externes** : alimenter le formulaire depuis un service web, un fichier ou une autre application, là où il n'existe pas de table à lier.
- **Maîtrise du multi-utilisateur** : aucune liaison permanente au back-end, donc moins de verrous et de trafic réseau (chapitre 15).
- **Contrôle total du flux** : décider précisément quand et comment l'enregistrement, l'annulation ou la réinitialisation surviennent.

Les formulaires de recherche relèvent aussi de cette logique ; ils sont traités plus spécifiquement au chapitre 6.5.

---

## 6.8.3. Contreparties : ce que l'on perd

Un formulaire indépendant abandonne tous les automatismes du mode lié, qu'il faut donc réécrire :

- la **navigation** entre enregistrements ;
- l'**enregistrement automatique** au changement d'enregistrement ou à la fermeture ;
- le **suivi des modifications** (propriété `Dirty`) ;
- la **validation** au niveau du formulaire ;
- le **verrouillage** des enregistrements en multi-utilisateur.

> **Conséquence sur les événements** : les événements de contrôle (`Change`, `AfterUpdate` d'un contrôle) se déclenchent toujours, mais les événements d'**enregistrement** du formulaire (`Current`, `BeforeUpdate`, `AfterUpdate` au sens « sauvegarde d'un enregistrement ») ne se produisent pas, faute d'enregistrement lié. Ces événements sont détaillés au chapitre 8.8.

Le choix entre lié et indépendant est donc un arbitrage : simplicité contre contrôle.

---

## 6.8.4. Le cycle de vie : charger → modifier → enregistrer

Un formulaire indépendant s'organise autour de trois opérations explicites, qu'il est recommandé d'isoler dans des procédures distinctes :

1. **Charger** un enregistrement : lire les données et les déverser dans les contrôles.
2. **Modifier** en mémoire : l'utilisateur édite librement les contrôles, sans impact sur la base.
3. **Enregistrer** : écrire le contenu des contrôles vers la base, en distinguant création et modification.

Une variable de module mémorise l'identifiant de l'enregistrement courant pour faire le lien entre ces étapes :

```vba
' En tête du module du formulaire
Option Compare Database
Option Explicit

Private m_lngClientID As Long   ' 0 = nouvel enregistrement, sinon ID en modification
Private m_bModifie    As Boolean ' indicateur de modification (dirty manuel)
```

---

## 6.8.5. Charger des données dans les contrôles

Le chargement ouvre un `Recordset` sur l'enregistrement voulu et recopie chaque champ dans le contrôle correspondant. La fonction `Nz` protège contre les valeurs `Null`.

```vba
Private Sub ChargerEnregistrement(ByVal lngID As Long)
    Dim rs As DAO.Recordset
    Dim strSQL As String

    strSQL = "SELECT * FROM T_Clients WHERE ClientID = " & lngID
    Set rs = CurrentDb.OpenRecordset(strSQL, dbOpenSnapshot)

    If Not rs.EOF Then
        Me.txtNom.Value   = Nz(rs!Nom, "")
        Me.txtVille.Value = Nz(rs!Ville, "")
        Me.txtEmail.Value = Nz(rs!Email, "")
        m_lngClientID = lngID
    End If

    rs.Close
    Set rs = Nothing
    m_bModifie = False   ' on vient de charger : aucune modification en attente
End Sub
```

Les types de `Recordset` (`dbOpenSnapshot`, `dbOpenDynaset`…) et la lecture des champs sont détaillés aux chapitres 9.3 et 9.5.

---

## 6.8.6. Enregistrer : la voie Recordset DAO (recommandée)

Deux techniques permettent d'écrire les données : construire une instruction SQL `INSERT`/`UPDATE`, ou passer par un `Recordset` DAO en modification. **La seconde est généralement préférable** pour un formulaire indépendant : elle évite la construction manuelle de SQL et donc tous les pièges d'échappement des chaînes, de formatage des dates et de séparateur décimal (voir chapitre 6.5).

```vba
Private Sub EnregistrerEnregistrement()
    Dim rs As DAO.Recordset

    If m_lngClientID = 0 Then
        ' --- Création ---
        Set rs = CurrentDb.OpenRecordset("T_Clients", dbOpenDynaset, dbAppendOnly)
        rs.AddNew
    Else
        ' --- Modification ---
        Set rs = CurrentDb.OpenRecordset( _
            "SELECT * FROM T_Clients WHERE ClientID = " & m_lngClientID, _
            dbOpenDynaset)
        rs.Edit
    End If

    rs!Nom   = Me.txtNom.Value
    rs!Ville = Me.txtVille.Value
    rs!Email = Me.txtEmail.Value
    rs.Update

    ' Récupérer la clé du nouvel enregistrement (NuméroAuto)
    If m_lngClientID = 0 Then
        rs.Bookmark = rs.LastModified
        m_lngClientID = rs!ClientID
    End If

    rs.Close
    Set rs = Nothing
    m_bModifie = False
End Sub
```

Le couple `rs.Bookmark = rs.LastModified` suivi de la lecture de la clé est la technique fiable pour récupérer la valeur d'un champ NuméroAuto fraîchement créé. L'ajout, la modification et la lecture des enregistrements en DAO sont approfondis aux chapitres 9.6 et 9.7.

> **Si l'on tient à passer par SQL** (`CurrentDb.Execute`), il devient impératif de paramétrer la requête pour échapper correctement les valeurs et prévenir l'injection SQL (chapitre 11.5), et de gérer le formatage localisé des dates et nombres (chapitre 11.7).

---

## 6.8.7. Suivre l'état : nouvel enregistrement ou modification

La variable de module `m_lngClientID` matérialise l'état courant :

- `m_lngClientID = 0` → le formulaire prépare une **création** ;
- `m_lngClientID > 0` → le formulaire édite l'enregistrement correspondant.

La routine de réinitialisation prépare une nouvelle saisie en vidant les contrôles et en remettant l'état à zéro :

```vba
Private Sub ViderControles()
    Me.txtNom.Value   = Null
    Me.txtVille.Value = Null
    Me.txtEmail.Value = Null
    m_lngClientID = 0
    m_bModifie = False
End Sub
```

---

## 6.8.8. Validation manuelle des données

En l'absence de validation automatique, on regroupe les contrôles de cohérence dans une fonction appelée **avant** l'enregistrement. Elle renvoie `True` si la saisie est valide, et signale sinon le premier problème en repositionnant le focus.

```vba
Private Function DonneesValides() As Boolean
    DonneesValides = False

    If Len(Nz(Me.txtNom.Value, "")) = 0 Then
        MsgBox "Le nom est obligatoire.", vbExclamation, "Validation"
        Me.txtNom.SetFocus
        Exit Function
    End If

    If Len(Nz(Me.txtEmail.Value, "")) > 0 _
       And InStr(Me.txtEmail.Value, "@") = 0 Then
        MsgBox "L'adresse e-mail est invalide.", vbExclamation, "Validation"
        Me.txtEmail.SetFocus
        Exit Function
    End If

    DonneesValides = True
End Function
```

Le bouton d'enregistrement enchaîne validation puis écriture :

```vba
Private Sub cmdEnregistrer_Click()
    If Not DonneesValides() Then Exit Sub
    EnregistrerEnregistrement
    MsgBox "Enregistrement effectué.", vbInformation
End Sub
```

Les techniques de validation et de messages d'erreur personnalisés sont développées au chapitre 8.10.

---

## 6.8.9. Détecter les modifications (dirty tracking manuel)

Pour avertir l'utilisateur avant qu'il ne perde une saisie non enregistrée, on maintient soi-même un indicateur de modification. Il passe à `True` dès qu'un contrôle change, et l'on l'interroge avant de charger un autre enregistrement ou de fermer le formulaire.

```vba
' Sur l'événement Change (ou AfterUpdate) de chaque contrôle de saisie
Private Sub txtNom_Change()
    m_bModifie = True
End Sub

Private Sub txtVille_Change()
    m_bModifie = True
End Sub
```

```vba
' Avant de quitter ou de charger autre chose
Private Sub Form_Unload(Cancel As Integer)
    If m_bModifie Then
        If MsgBox("Des modifications n'ont pas été enregistrées. Quitter quand même ?", _
                  vbYesNo + vbQuestion, "Confirmation") = vbNo Then
            Cancel = True
        End If
    End If
End Sub
```

Ce mécanisme reproduit, à la main, ce que la propriété `Dirty` offre nativement à un formulaire lié.

---

## 6.8.10. Pièges et bonnes pratiques

- **Isoler les opérations** dans des procédures dédiées (`ChargerEnregistrement`, `EnregistrerEnregistrement`, `ViderControles`, `DonneesValides`) : la lisibilité et la maintenance y gagnent considérablement.
- **Toujours protéger la lecture des contrôles avec `Nz`** : un contrôle vide renvoie `Null`, ce qui propage des erreurs dans les concaténations et affectations.
- **Privilégier le `Recordset` DAO à la construction de SQL** pour écrire : on s'épargne l'échappement des apostrophes, le format des dates et le séparateur décimal.
- **Si l'on construit du SQL malgré tout**, paramétrer la requête (chapitre 11.5) et gérer la localisation (chapitre 11.7).
- **Réinitialiser systématiquement l'état** (`m_lngClientID`, `m_bModifie`) entre deux enregistrements, sous peine de mélanger création et modification.
- **Récupérer la clé après insertion** via `rs.Bookmark = rs.LastModified` pour enchaîner sur la saisie de données liées (lignes de détail, par exemple).
- **Penser à la transaction** lorsqu'une saisie écrit dans plusieurs tables : encadrer les écritures par une transaction pour garantir l'atomicité (chapitre 14).
- **Structurer l'application** : pour des saisies complexes récurrentes, les patrons Repository (chapitre 16.6) et l'architecture en couches (chapitre 16.8) clarifient la séparation entre interface et accès aux données.

---

## 6.8.11. Récapitulatif

- Un formulaire indépendant n'a ni `RecordSource` ni `ControlSource` : le développeur pilote tout par code, en échange d'un contrôle total.
- Il s'impose pour les saisies réparties sur plusieurs tables, la validation avant écriture, les assistants multi-étapes et les sources de données externes.
- Son cycle de vie repose sur trois opérations explicites — **charger**, **modifier**, **enregistrer** — qu'il convient d'isoler dans des procédures distinctes, avec une variable d'état pour la clé courante.
- L'enregistrement par **`Recordset` DAO** (`AddNew`/`Edit` + `Update`) est préférable à la construction de SQL : il évite les pièges d'échappement et de formatage.
- En l'absence d'automatismes, il faut réimplémenter la **validation**, le **suivi des modifications** (`Dirty` manuel) et, si nécessaire, la navigation.
- Pour des écritures multi-tables, encadrer les opérations par une **transaction** (chapitre 14) ; pour des saisies récurrentes, s'appuyer sur les patrons d'architecture (chapitres 16.6 et 16.8).

⏭️ [6.9. Communication entre formulaires (sans couplage fort)](/06-formulaires/09-communication-formulaires.md)
