🔝 Retour au [Sommaire](/SOMMAIRE.md)

# K. Modèles de code réutilisables (snippets)

Cette annexe rassemble des blocs de code prêts à adapter, illustrant les patterns récurrents de la formation. Chaque modèle est volontairement compact ; le chapitre indiqué entre parenthèses en développe le contexte et les variantes.

> ℹ️ Conventions communes : `Option Explicit` en tête de module, gestion d'erreurs avec étiquettes `Sortie` / `GestionErreur`, **liaison tardive** (`CreateObject`) pour la portabilité entre versions, et **libération systématique** des objets (`Set obj = Nothing`). Adapter les noms d'objets et chemins entre `< >` ou en exemple.

---

## Gestion des erreurs

### Squelette de procédure avec gestion d'erreurs (ch. 13.2)

```vba
Public Sub MaProcedure()
    On Error GoTo GestionErreur

    ' --- Code principal ---

Sortie:
    ' --- Nettoyage, toujours exécuté ---
    Exit Sub

GestionErreur:
    MsgBox "Erreur " & Err.Number & " : " & Err.Description, _
           vbExclamation, "MaProcedure"
    Resume Sortie
End Sub
```

### Journalisation d'une erreur dans une table (ch. 13.6)

*Utilise le modèle `SQLTexte` plus bas. La journalisation ne doit jamais interrompre l'application.*

```vba
Public Sub JournaliserErreur(ByVal numero As Long, _
                             ByVal description As String, _
                             ByVal procedure As String)
    On Error Resume Next
    CurrentDb.Execute _
        "INSERT INTO Journal_Erreurs (DateErreur, Numero, Description, Procedure) " & _
        "VALUES (Now(), " & numero & ", " & SQLTexte(description) & ", " & _
        SQLTexte(procedure) & ")", dbFailOnError
End Sub
```

---

## Accès aux données — DAO

### Parcours d'un Recordset (ch. 9.3, 9.4, 9.11)

```vba
Public Sub ParcourirClients()
    Dim db As DAO.Database
    Dim rst As DAO.Recordset

    On Error GoTo GestionErreur
    Set db = CurrentDb
    Set rst = db.OpenRecordset("SELECT * FROM Clients", dbOpenSnapshot)

    If Not (rst.EOF And rst.BOF) Then
        Do Until rst.EOF
            Debug.Print rst!Nom
            rst.MoveNext
        Loop
    End If

Sortie:
    If Not rst Is Nothing Then rst.Close
    Set rst = Nothing
    Set db = Nothing
    Exit Sub

GestionErreur:
    MsgBox "Erreur " & Err.Number & " : " & Err.Description, vbExclamation
    Resume Sortie
End Sub
```

### Ajout et modification d'un enregistrement (ch. 9.6, 9.7)

```vba
' Ajout
With rst
    .AddNew
    !Nom = "Durand"
    !Ville = "Rouen"
    .Update
End With

' Modification
With rst
    .Edit
    !Solde = !Solde + 100
    .Update
End With
```

### Exécuter une requête action en toute sécurité (ch. 11.1, 18.3)

*À préférer à `DoCmd.RunSQL` : aucun message de confirmation, et les erreurs sont levées (`dbFailOnError`).*

```vba
CurrentDb.Execute _
    "UPDATE Clients SET Actif = False WHERE DerniereCommande < #2025/01/01#", _
    dbFailOnError
```

### Requête paramétrée (prévention des injections) (ch. 11.5, 12.3)

```vba
Dim qdf As DAO.QueryDef
Dim rst As DAO.Recordset

Set qdf = CurrentDb.CreateQueryDef("")    ' requête temporaire
qdf.SQL = "PARAMETERS pVille TEXT(50); " & _
          "SELECT * FROM Clients WHERE Ville = [pVille]"
qdf.Parameters("pVille").Value = Me.txtVille
Set rst = qdf.OpenRecordset(dbOpenSnapshot)
```

---

## Accès aux données — ADO

### Connexion et Recordset (liaison tardive) (ch. 10.3, 10.4)

```vba
Dim cn As Object, rs As Object

Set cn = CreateObject("ADODB.Connection")
cn.Open "Provider=MSOLEDBSQL;Server=SRV1;Database=GestCom;Trusted_Connection=yes;"

Set rs = CreateObject("ADODB.Recordset")
rs.Open "SELECT * FROM dbo.Clients", cn, 3, 1   ' adOpenStatic, adLockReadOnly

Do Until rs.EOF
    Debug.Print rs.Fields("Nom").Value
    rs.MoveNext
Loop

rs.Close: Set rs = Nothing
cn.Close: Set cn = Nothing
```

---

## Transactions

### Transaction DAO (ch. 14.2)

```vba
Public Sub TransfertFonds()
    Dim ws As DAO.Workspace
    Dim db As DAO.Database
    Dim transEnCours As Boolean

    On Error GoTo GestionErreur
    Set ws = DBEngine(0)
    Set db = CurrentDb

    ws.BeginTrans
    transEnCours = True

    db.Execute "UPDATE Comptes SET Solde = Solde - 100 WHERE ID = 1", dbFailOnError
    db.Execute "UPDATE Comptes SET Solde = Solde + 100 WHERE ID = 2", dbFailOnError

    ws.CommitTrans
    transEnCours = False

Sortie:
    Set db = Nothing
    Set ws = Nothing
    Exit Sub

GestionErreur:
    If transEnCours Then ws.Rollback
    MsgBox "Transaction annulée : " & Err.Description, vbCritical
    Resume Sortie
End Sub
```

---

## Formulaires

### Vérifier qu'un formulaire est ouvert (ch. 6.2)

```vba
Public Function FormulaireOuvert(ByVal nomForm As String) As Boolean
    FormulaireOuvert = _
        (SysCmd(acSysCmdGetObjectState, acForm, nomForm) And acObjStateOpen) <> 0
End Function
```

### Localiser un enregistrement via RecordsetClone (ch. 9.12)

```vba
With Me.RecordsetClone
    .FindFirst "ID = " & Me.txtRechercheID
    If Not .NoMatch Then Me.Bookmark = .Bookmark
End With
```

### DLookup sécurisé (ch. 11.10)

```vba
Dim sNom As String
sNom = Nz(DLookup("Nom", "Clients", "ID = " & lngID), "(inconnu)")
```

---

## Utilitaires SQL dynamique

### Encadrer une chaîne pour SQL (échappement des apostrophes) (ch. 11.6)

```vba
Public Function SQLTexte(ByVal s As String) As String
    SQLTexte = "'" & Replace(s, "'", "''") & "'"
End Function
```

### Formater une date pour SQL (format ISO, indépendant de la locale) (ch. 11.7)

```vba
Public Function SQLDate(ByVal d As Date) As String
    SQLDate = "#" & Format$(d, "yyyy\/mm\/dd") & "#"
End Function
```

---

## Interface et performance

### Encadrer un traitement long (ch. 5.7, 18)

```vba
Public Sub TraitementLong()
    On Error GoTo Fin
    DoCmd.Hourglass True
    DoCmd.Echo False
    DoCmd.SetWarnings False

    ' --- traitement ---

Fin:
    DoCmd.SetWarnings True
    DoCmd.Echo True
    DoCmd.Hourglass False
    If Err.Number <> 0 Then MsgBox Err.Description, vbExclamation
End Sub
```

### Barre de progression dans la barre d'état (ch. 17.5)

```vba
Dim varRet As Variant
varRet = SysCmd(acSysCmdInitMeter, "Traitement en cours…", lngTotal)
' Dans la boucle :
SysCmd acSysCmdUpdateMeter, i
' À la fin :
SysCmd acSysCmdRemoveMeter
```

---

## Collections

### Scripting.Dictionary (ch. 3.7)

```vba
Dim dico As Object
Set dico = CreateObject("Scripting.Dictionary")

dico("cle1") = "valeur1"
If dico.Exists("cle1") Then Debug.Print dico("cle1")

Dim k As Variant
For Each k In dico.Keys
    Debug.Print k, dico(k)
Next k
```

---

## Système et fichiers

### Vérifier l'existence d'un fichier (ch. 22.10)

```vba
If Len(Dir$("C:\Data\base.accdb")) > 0 Then
    ' Le fichier existe
End If
```

### Lire et écrire un fichier texte (FileSystemObject) (ch. 22.10)

```vba
Dim fso As Object, ts As Object
Set fso = CreateObject("Scripting.FileSystemObject")

' Écriture
Set ts = fso.CreateTextFile("C:\Data\log.txt", True)
ts.WriteLine "Ligne de journal"
ts.Close

' Lecture
Set ts = fso.OpenTextFile("C:\Data\log.txt", 1)   ' ForReading
Do Until ts.AtEndOfStream
    Debug.Print ts.ReadLine
Loop
ts.Close
Set ts = Nothing: Set fso = Nothing
```

---

## Intégration Office

### Envoyer un e-mail via Outlook (liaison tardive) (ch. 22.5)

```vba
Dim ol As Object, msg As Object
Set ol = CreateObject("Outlook.Application")
Set msg = ol.CreateItem(0)              ' olMailItem

With msg
    .To = "destinataire@example.com"
    .Subject = "Sujet"
    .Body = "Corps du message"
    .Display                            ' ou .Send
End With

Set msg = Nothing: Set ol = Nothing
```

### Exporter une table vers Excel (ch. 22.3)

```vba
DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel12Xml, _
    "Clients", "C:\Export\Clients.xlsx", True
```

---

## Programmation orientée objet

### Squelette de module de classe avec propriété (ch. 16.1, 16.2)

*Dans un module de classe nommé `Client`.*

```vba
Option Explicit

Private mNom As String

Public Property Get Nom() As String
    Nom = mNom
End Property

Public Property Let Nom(ByVal valeur As String)
    mNom = valeur
End Property
```

Instanciation et usage depuis un module standard :

```vba
Dim oClient As Client
Set oClient = New Client
oClient.Nom = "Durand"
Debug.Print oClient.Nom
Set oClient = Nothing
```

---

> 💡 **À retenir.** Ces modèles partagent une même rigueur : **gestion d'erreurs systématique**, **libération des objets**, **paramétrage** plutôt que concaténation directe dans le SQL, et `Execute … dbFailOnError` plutôt que `RunSQL`. Combinés aux conventions de nommage de l'**annexe G** et aux bonnes pratiques du **chapitre 24**, ils constituent une base saine pour démarrer rapidement sans réécrire les fondations à chaque projet.

⏭️ [Retour au Sommaire](/SOMMAIRE.md)
