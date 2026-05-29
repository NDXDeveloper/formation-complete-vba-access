🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 16.8. Architecture en couches (DAL, BLL, UI) appliquée à Access

L'architecture en couches est l'aboutissement naturel des notions abordées dans ce chapitre. Les modules de classe (16.1), les propriétés (16.2), l'encapsulation (16.3), le patron Repository (16.6) et la Factory (16.7) ne sont pas des techniques isolées : ce sont les briques qui permettent d'organiser une application Access selon une structure claire, durable et testable. Cette section montre comment assembler ces briques en trois couches — **DAL** (accès aux données), **BLL** (logique métier) et **UI** (présentation) — et, surtout, comment adapter ce modèle aux contraintes propres d'Access, notamment l'omniprésence des formulaires liés.

## Pourquoi structurer une application Access en couches

Une base Access qui démarre comme un petit outil finit presque toujours par grossir : on ajoute des formulaires, des règles de gestion, des contrôles de cohérence, des exports. Sans organisation, ces ajouts s'empilent dans le code-behind des formulaires, là où ils sont le plus rapides à écrire. Au bout de quelques mois, le moindre changement de règle métier oblige à fouiller dans une dizaine de modules de formulaire, et la même validation est dupliquée à trois endroits avec de subtiles différences.

L'architecture en couches répond à ce problème par le principe de **séparation des responsabilités** (*separation of concerns*) : chaque partie du code a un rôle unique et bien délimité. La présentation s'occupe de l'affichage et des interactions. La logique métier porte les règles de l'application. L'accès aux données encapsule tout ce qui touche à la base. Une règle de calcul de remise se trouve à un seul endroit ; une modification de la structure des tables n'impacte qu'une couche ; et l'on peut faire évoluer l'interface sans toucher au cœur métier.

## Le piège du code « tout dans le formulaire »

L'anti-modèle le plus répandu en Access consiste à concentrer la totalité de l'application dans le code-behind des formulaires. Voici un bouton d'enregistrement typique d'une application non structurée :

```vba
' ANTI-MODÈLE : tout est mélangé dans le formulaire
Private Sub cmdEnregistrer_Click()
    Dim db As DAO.Database
    Dim rs As DAO.Recordset

    ' Validation métier directement dans l'UI
    If Len(Trim$(Me!txtNom & "")) = 0 Then
        MsgBox "Le nom est obligatoire."
        Exit Sub
    End If
    If InStr(Me!txtEmail & "", "@") = 0 Then
        MsgBox "Email invalide."
        Exit Sub
    End If

    ' Accès aux données directement dans l'UI
    Set db = CurrentDb
    Set rs = db.OpenRecordset("tblClients", dbOpenDynaset, dbAppendOnly)
    rs.AddNew
    rs!Nom = Me!txtNom
    rs!Email = Me!txtEmail
    rs!DateInscription = Now()
    rs.Update
    rs.Close
    Set rs = Nothing
    Set db = Nothing

    MsgBox "Enregistré."
End Sub
```

Ce code fonctionne, mais il mélange trois préoccupations distinctes : la réaction à un clic (UI), les règles de validation (métier) et l'écriture en base (données). Les conséquences se paient plus tard. La règle « l'email doit être unique » devra être recopiée dans chaque formulaire qui crée un client. Une migration vers SQL Server (chapitre 23) imposera de modifier chaque procédure de chaque formulaire. Et il devient quasiment impossible de tester la logique métier sans ouvrir le formulaire et cliquer manuellement.

## Les trois couches et leurs responsabilités

L'architecture en couches répartit ce même code en trois ensembles aux rôles strictement délimités.

La **couche de présentation (UI)** regroupe les formulaires, les états et les contrôles. Son unique rôle est de présenter des informations à l'utilisateur et de capter ses actions. Elle ne contient aucune règle métier et n'accède jamais directement à la base. Quand l'utilisateur agit, elle se contente de transmettre la demande à la couche inférieure et d'afficher le résultat (ou l'erreur) qui lui revient.

La **couche de logique métier (BLL — *Business Logic Layer*)** porte les règles de l'application : validations, calculs, contrôles de cohérence, orchestration des opérations. C'est ici que vit la connaissance du domaine (« un client ne peut pas être supprimé s'il a des commandes en cours », « la remise maximale est de 30 % »). Elle ne sait rien de l'interface (elle ignore l'existence des formulaires et des contrôles) et ne connaît pas non plus les détails techniques de la base : elle délègue toute persistance à la couche inférieure.

La **couche d'accès aux données (DAL — *Data Access Layer*)** est la seule à connaître DAO, ADO, le SQL et le nom des tables. Elle expose des opérations de haut niveau (`GetById`, `Insert`, `Delete`...) et masque entièrement la mécanique des `Recordset`. C'est, dans la pratique, le patron Repository vu en section 16.6, replacé dans une architecture globale.

| Couche | Connaît | Ne connaît PAS | Exemple de classe |
|--------|---------|----------------|-------------------|
| UI | Les formulaires, contrôles, événements | Le SQL, les tables, les règles métier | `Form_frmInscriptionClient` |
| BLL | Les règles métier, le domaine | Les formulaires, DAO/ADO, le SQL | `clsClientService` |
| DAL | DAO/ADO, le SQL, les tables | Les règles métier, les formulaires | `clsClientDAL` |

## La règle de dépendance : un sens unique

Le principe le plus important de toute architecture en couches est la **direction des dépendances**. Les couches ne se connaissent que vers le bas : l'UI appelle la BLL, la BLL appelle la DAL, la DAL lit et écrit en base. Jamais l'inverse. La DAL n'appelle jamais la BLL, et la BLL ne manipule jamais un formulaire.

```
┌──────────────────────────────────────────┐
│   UI — Présentation                        │
│   Formulaires, états, contrôles            │
│   (Form_frmInscriptionClient)              │
└────────────────────┬───────────────────────┘
                     │ appelle
                     ▼
┌──────────────────────────────────────────┐
│   BLL — Logique métier                     │
│   Règles, validations, orchestration       │
│   (clsClientService)                       │
└────────────────────┬───────────────────────┘
                     │ appelle
                     ▼
┌──────────────────────────────────────────┐
│   DAL — Accès aux données                  │
│   CRUD, SQL, DAO/ADO                        │
│   (clsClientDAL)                           │
└────────────────────┬───────────────────────┘
                     │ lit / écrit
                     ▼
┌──────────────────────────────────────────┐
│   Base de données (tables)                 │
│   tblClients                               │
└──────────────────────────────────────────┘

   L'entité clsClient circule entre toutes les couches
```

Ce sens unique a une conséquence pratique précieuse : on peut remplacer une couche supérieure sans toucher aux couches inférieures. Changer entièrement l'interface (passer de formulaires liés à des formulaires indépendants, par exemple) n'affecte ni la BLL ni la DAL. À l'inverse, comme les couches basses ignorent les couches hautes, elles restent réutilisables : la même DAL peut servir un formulaire, un état, ou une routine d'import nocturne sans interface du tout.

## L'entité de domaine : le contrat qui circule entre les couches

Pour que les couches s'échangent des données proprement, on évite de faire transiter des `Recordset` ou des valeurs éparses entre elles. On définit à la place une **entité de domaine** : un module de classe simple qui représente un objet métier. Ici, un client.

```vba
' ============================================
' Module de classe : clsClient
' Couche : Entité / Modèle de domaine
' Objet métier pur, partagé par toutes les couches.
' Ne contient ni accès aux données, ni logique de
' présentation : uniquement des données et,
' éventuellement, des règles internes à l'entité.
' ============================================
Option Compare Database
Option Explicit

Private m_ID As Long
Private m_Nom As String
Private m_Email As String
Private m_DateInscription As Date
Private m_Actif As Boolean

Public Property Get ID() As Long
    ID = m_ID
End Property
Public Property Let ID(ByVal Value As Long)
    m_ID = Value
End Property

Public Property Get Nom() As String
    Nom = m_Nom
End Property
Public Property Let Nom(ByVal Value As String)
    m_Nom = Value
End Property

Public Property Get Email() As String
    Email = m_Email
End Property
Public Property Let Email(ByVal Value As String)
    m_Email = Value
End Property

Public Property Get DateInscription() As Date
    DateInscription = m_DateInscription
End Property
Public Property Let DateInscription(ByVal Value As Date)
    m_DateInscription = Value
End Property

Public Property Get Actif() As Boolean
    Actif = m_Actif
End Property
Public Property Let Actif(ByVal Value As Boolean)
    m_Actif = Value
End Property

' Un nouveau client (non encore enregistré) a un ID à zéro.
Public Property Get EstNouveau() As Boolean
    EstNouveau = (m_ID = 0)
End Property
```

Cette entité est le **contrat partagé** : la DAL la produit (en lisant une ligne) et la consomme (en écrivant une ligne), la BLL la manipule, et l'UI affiche ses propriétés. Comme elle ne dépend d'aucune technologie, elle traverse toutes les couches sans les coupler entre elles.

## La couche d'accès aux données (DAL)

La DAL encapsule tout l'accès à `tblClients`. C'est la seule classe du système qui mentionne DAO, le SQL ou le nom des champs. Elle reçoit et renvoie des `clsClient`, jamais des `Recordset`.

```vba
' ============================================
' Module de classe : clsClientDAL
' Couche : Accès aux données (DAL)
' Seule couche qui connaît DAO, le SQL et les tables.
' Reçoit/renvoie des clsClient ; ne fait remonter
' aucun Recordset vers les couches supérieures.
' ============================================
Option Compare Database
Option Explicit

Private Const TABLE_NAME As String = "tblClients"

' --- Lecture par identifiant ---
Public Function GetById(ByVal clientId As Long) As clsClient
    Dim db As DAO.Database
    Dim rs As DAO.Recordset

    Set db = CurrentDb
    ' clientId est numérique : pas de risque d'injection ici.
    ' Pour tout paramètre de type texte, préférer une requête
    ' paramétrée (voir section 11.5).
    Set rs = db.OpenRecordset( _
        "SELECT * FROM " & TABLE_NAME & " WHERE ID = " & clientId, _
        dbOpenSnapshot)

    If Not rs.EOF Then
        Set GetById = MapRecordToClient(rs)
    End If

    rs.Close
    Set rs = Nothing
    Set db = Nothing
End Function

' --- Vérification d'unicité d'un email ---
Public Function EmailExiste(ByVal email As String) As Boolean
    Dim db As DAO.Database
    Dim qd As DAO.QueryDef
    Dim rs As DAO.Recordset

    Set db = CurrentDb
    ' Requête paramétrée : robuste face aux apostrophes et injections.
    Set qd = db.CreateQueryDef("", _
        "SELECT Count(*) AS Nb FROM " & TABLE_NAME & " WHERE Email = [pEmail]")
    qd.Parameters("pEmail") = email
    Set rs = qd.OpenRecordset(dbOpenSnapshot)

    EmailExiste = (rs!Nb > 0)

    rs.Close
    Set rs = Nothing
    Set qd = Nothing
    Set db = Nothing
End Function

' --- Insertion ; renvoie l'ID auto-généré ---
Public Function Insert(ByVal client As clsClient) As Long
    Dim db As DAO.Database
    Dim rs As DAO.Recordset

    Set db = CurrentDb
    Set rs = db.OpenRecordset(TABLE_NAME, dbOpenDynaset, dbAppendOnly)

    rs.AddNew
    rs!Nom = client.Nom
    rs!Email = client.Email
    rs!DateInscription = client.DateInscription
    rs!Actif = client.Actif
    rs.Update

    ' Pour lire le NuméroAuto généré, on se positionne sur
    ' le dernier enregistrement modifié.
    rs.Bookmark = rs.LastModified
    Insert = rs!ID

    rs.Close
    Set rs = Nothing
    Set db = Nothing
End Function

' --- Mise à jour d'un client existant ---
Public Sub Update(ByVal client As clsClient)
    Dim db As DAO.Database
    Dim rs As DAO.Recordset

    Set db = CurrentDb
    Set rs = db.OpenRecordset( _
        "SELECT * FROM " & TABLE_NAME & " WHERE ID = " & client.ID, _
        dbOpenDynaset)

    If Not rs.EOF Then
        rs.Edit
        rs!Nom = client.Nom
        rs!Email = client.Email
        rs!Actif = client.Actif
        rs.Update
    End If

    rs.Close
    Set rs = Nothing
    Set db = Nothing
End Sub

' --- Conversion d'une ligne en entité (mapping) ---
Private Function MapRecordToClient(ByRef rs As DAO.Recordset) As clsClient
    Dim c As clsClient
    Set c = New clsClient

    c.ID = rs!ID
    c.Nom = Nz(rs!Nom, "")
    c.Email = Nz(rs!Email, "")
    c.DateInscription = Nz(rs!DateInscription, #1/1/1900#)
    c.Actif = Nz(rs!Actif, False)

    Set MapRecordToClient = c
End Function
```

Le point clé est la fonction privée `MapRecordToClient` : elle traduit une ligne de table en objet métier. C'est la frontière étanche entre le monde des tables et le monde des objets. À partir du moment où une donnée quitte la DAL, elle a la forme d'un `clsClient` ; les couches supérieures n'ont jamais à connaître l'existence d'un champ nommé `DateInscription`.

## La couche de logique métier (BLL)

La BLL applique les règles de l'application et orchestre la DAL. Elle ne touche jamais à un formulaire ni à un `Recordset` : elle ne fait que vérifier des règles, construire des entités et appeler la DAL pour les persister.

```vba
' ============================================
' Module de classe : clsClientService
' Couche : Logique métier (BLL)
' Applique les règles métier et orchestre la DAL.
' Ne connaît NI les formulaires NI DAO/SQL.
' ============================================
Option Compare Database
Option Explicit

Private m_dal As clsClientDAL

Private Sub Class_Initialize()
    Set m_dal = New clsClientDAL
End Sub

Private Sub Class_Terminate()
    Set m_dal = Nothing
End Sub

' Inscrit un nouveau client après validation des règles métier.
' Renvoie l'identifiant du client créé.
Public Function InscrireClient(ByVal nom As String, _
                               ByVal email As String) As Long
    Dim client As clsClient

    ' --- Règles métier (centralisées ici, une seule fois) ---
    If Len(Trim$(nom)) = 0 Then
        Err.Raise vbObjectError + 1001, "clsClientService.InscrireClient", _
                  "Le nom du client est obligatoire."
    End If

    If Not EmailValide(email) Then
        Err.Raise vbObjectError + 1002, "clsClientService.InscrireClient", _
                  "L'adresse email saisie n'est pas valide."
    End If

    If m_dal.EmailExiste(LCase$(Trim$(email))) Then
        Err.Raise vbObjectError + 1003, "clsClientService.InscrireClient", _
                  "Un client utilise déjà cette adresse email."
    End If

    ' --- Construction de l'entité ---
    Set client = New clsClient
    client.Nom = Trim$(nom)
    client.Email = LCase$(Trim$(email))
    client.DateInscription = Now()
    client.Actif = True

    ' --- Persistance déléguée à la DAL ---
    InscrireClient = m_dal.Insert(client)
End Function

' Désactive un client (règle métier : on ne supprime jamais,
' on désactive, pour préserver l'historique).
Public Sub DesactiverClient(ByVal clientId As Long)
    Dim client As clsClient

    Set client = m_dal.GetById(clientId)
    If client Is Nothing Then
        Err.Raise vbObjectError + 1004, "clsClientService.DesactiverClient", _
                  "Client introuvable (ID " & clientId & ")."
    End If

    client.Actif = False
    m_dal.Update client
End Sub

' Validation interne d'un email (règle métier privée).
Private Function EmailValide(ByVal email As String) As Boolean
    Dim posArobase As Long
    posArobase = InStr(email, "@")
    EmailValide = (posArobase > 1) And _
                  (InStr(posArobase, email, ".") > posArobase + 1)
End Function
```

La gestion des erreurs mérite attention. La BLL signale les violations de règles métier en levant des erreurs personnalisées avec `Err.Raise` (technique détaillée en section 13.8). Elle ne décide pas de la manière de les afficher — c'est le rôle de l'UI. Ce partage est essentiel : la même règle pourra ainsi remonter une erreur affichée dans une `MsgBox` depuis un formulaire, ou journalisée dans un fichier depuis un traitement automatisé, sans rien changer à la BLL. La hiérarchie de propagation des erreurs entre procédures est traitée en section 13.9.

## La couche de présentation (UI)

Le code-behind du formulaire devient remarquablement maigre. Il capte le clic, transmet les valeurs à la BLL et affiche le résultat. Aucune règle, aucun SQL.

```vba
' ============================================
' Code-behind du formulaire frmInscriptionClient
' Couche : Présentation (UI)
' Aucune règle métier, aucun accès aux données.
' Délègue tout à la couche métier (BLL).
' ============================================
Option Compare Database
Option Explicit

Private Sub cmdEnregistrer_Click()
    Dim service As clsClientService
    Dim nouvelId As Long

    On Error GoTo GestionErreur

    Set service = New clsClientService
    nouvelId = service.InscrireClient(Me!txtNom & "", Me!txtEmail & "")

    MsgBox "Client enregistré avec succès (ID " & nouvelId & ").", _
           vbInformation, "Inscription"

    ' Réinitialisation du formulaire pour une nouvelle saisie
    Me!txtNom = Null
    Me!txtEmail = Null
    Me!txtNom.SetFocus

SortiePropre:
    Set service = Nothing
    Exit Sub

GestionErreur:
    ' L'UII se contente d'AFFICHER l'erreur métier remontée par la BLL.
    MsgBox Err.Description, vbExclamation, "Inscription impossible"
    Resume SortiePropre
End Sub
```

Comparé à l'anti-modèle du début, le bénéfice saute aux yeux : la règle d'unicité de l'email, la validation du format, le calcul de la date d'inscription — tout cela vit désormais dans la BLL, à un seul endroit. Un second formulaire de saisie réutiliserait exactement le même `clsClientService`, sans une ligne de logique dupliquée.

## Le flux complet, de bout en bout

Suivons une inscription du clic jusqu'à la table, pour voir les couches s'enchaîner :

1. L'utilisateur saisit un nom et un email, puis clique sur **Enregistrer**. L'événement `cmdEnregistrer_Click` (UI) se déclenche.
2. L'UI instancie un `clsClientService` (BLL) et appelle `InscrireClient`, en lui passant les valeurs brutes des contrôles.
3. La BLL valide : nom non vide, email bien formé, email non déjà utilisé (pour ce dernier contrôle, elle interroge la DAL via `EmailExiste`).
4. Si une règle est violée, la BLL lève une erreur ; l'UI la capte dans son `GestionErreur` et l'affiche. Le flux s'arrête là.
5. Sinon, la BLL construit un `clsClient`, fixe la date et le statut, puis appelle `m_dal.Insert`.
6. La DAL ouvre un `Recordset`, écrit la ligne, récupère le `NuméroAuto` généré et le renvoie.
7. L'ID remonte la BLL puis l'UI, qui affiche le message de succès et réinitialise le formulaire.

Chaque couche n'a vu que ce qui la concerne : l'UI n'a jamais touché à un `Recordset`, la BLL n'a jamais connu le nom du champ `DateInscription`, la DAL n'a jamais su qu'une règle d'unicité existait.

## Adapter le modèle aux formulaires liés d'Access

Voici la difficulté centrale, et la raison pour laquelle l'architecture en couches « pure » des manuels ne se transpose pas mécaniquement à Access. Access est construit autour des **formulaires liés** (*bound forms*) : un formulaire dont la propriété `RecordSource` pointe vers une table ou une requête lit et écrit les données automatiquement, sans une ligne de code. C'est la force d'Access en développement rapide — et c'est aussi une violation directe du principe de séparation, puisque le formulaire (UI) parle alors *directement* à la base, court-circuitant la BLL et la DAL.

Renoncer aux formulaires liés pour respecter à la lettre l'architecture en couches reviendrait à se priver de l'atout majeur d'Access et à réécrire à la main tout ce qu'il offre gratuitement. Ce n'est presque jamais le bon arbitrage. La voie pragmatique consiste à **ne pas choisir l'un ou l'autre, mais à doser selon la complexité** :

Pour la **saisie et la consultation simples** (un formulaire de gestion de catalogue, une grille de paramètres), on garde les formulaires liés. La séparation pure ne rapporterait ici que de la lourdeur, pour une logique quasi inexistante.

Pour la **validation métier**, même sur un formulaire lié, on route les contrôles par la BLL. L'astuce est de brancher l'événement `Form_BeforeUpdate` sur le service métier : le formulaire reste lié pour l'affichage et l'édition, mais aucune écriture ne passe sans avoir traversé les règles de la BLL.

```vba
' Sur un formulaire LIÉ : on conserve le confort de la liaison
' tout en faisant valider l'enregistrement par la BLL avant écriture.
Private Sub Form_BeforeUpdate(Cancel As Integer)
    Dim service As clsClientService

    On Error GoTo GestionErreur
    Set service = New clsClientService

    ' La BLL valide les valeurs courantes du formulaire.
    ' Si une règle échoue, elle lève une erreur et l'écriture
    ' est annulée.
    service.ValiderAvantEnregistrement Me!txtNom & "", Me!txtEmail & ""

    Set service = Nothing
    Exit Sub

GestionErreur:
    MsgBox Err.Description, vbExclamation, "Validation"
    Cancel = True              ' Annule l'enregistrement par Access
    Set service = Nothing
End Sub
```

L'usage de `Cancel = True` pour bloquer un enregistrement non valide relève des patterns d'annulation d'événements vus en section 8.12.

Pour les **opérations complexes** (un traitement qui touche plusieurs tables, applique des calculs, déclenche des effets de bord), on quitte le formulaire lié et l'on passe par un formulaire indépendant (*unbound form*, section 6.8) servant uniquement d'interface à la BLL, comme dans l'exemple d'inscription plus haut. Ces opérations bénéficient pleinement de l'architecture, car c'est là que la logique est dense et que la duplication coûte cher.

La règle de décision tient en une phrase : **plus une opération porte de logique métier, plus l'architecture en couches est rentable ; plus elle se réduit à du CRUD trivial, plus le formulaire lié natif d'Access se justifie.**

## Organiser les couches dans le projet VBA

VBA ne dispose pas d'espaces de noms (*namespaces*) ni de dossiers de projet : tous les modules de classe cohabitent à plat dans l'explorateur de projets (section 2.2). La séparation en couches doit donc être rendue visible par une **convention de nommage** rigoureuse, dans le prolongement des conventions Access vues en section 24.1 (préfixes Leszynski/Reddick, annexe G).

Une convention efficace consiste à préfixer chaque classe par sa couche :

| Préfixe | Couche | Exemples |
|---------|--------|----------|
| `cls...Entity` ou suffixe neutre | Entité de domaine | `clsClient`, `clsCommande` |
| `cls...DAL` | Accès aux données | `clsClientDAL`, `clsCommandeDAL` |
| `cls...Service` | Logique métier | `clsClientService`, `clsCommandeService` |
| `Form_...` / `Report_...` | Présentation | imposé par Access |

Les modules standard (non-classe) restent réservés aux utilitaires transversaux qui n'appartiennent à aucune couche métier : helpers de chaîne, de date, point d'entrée de l'application. Pour les services à instance unique (un service de configuration, un journal d'erreurs centralisé comme en section 13.7), le patron Singleton de la section 16.7 évite de recréer l'objet à chaque appel.

## Rendre la DAL interchangeable (vers SQL Server)

L'architecture en couches prend toute sa valeur le jour d'une migration. Si l'application grandit au point de dépasser les limites du moteur ACE (2 Go, concurrence — voir sections 18.9 et 15.10), la migration vers SQL Server (chapitre 23) devient nécessaire. Avec une DAL bien isolée, **seule cette couche change** : la BLL et l'UI restent intactes, puisqu'elles ne connaissent que des `clsClient`.

On peut aller plus loin et rendre l'implémentation interchangeable grâce aux interfaces (`Implements`, section 16.5). On définit d'abord une interface décrivant le contrat de la DAL :

```vba
' ============================================
' Module de classe : IClientRepository
' Interface (contrat) de la couche d'accès aux données.
' Ne contient que des signatures, aucune implémentation.
' ============================================
Option Compare Database
Option Explicit

Public Function GetById(ByVal clientId As Long) As clsClient
End Function

Public Function Insert(ByVal client As clsClient) As Long
End Function

Public Sub Update(ByVal client As clsClient)
End Sub

Public Function EmailExiste(ByVal email As String) As Boolean
End Function
```

La DAL concrète déclare implémenter ce contrat. Les méthodes deviennent privées et préfixées du nom de l'interface :

```vba
' Dans clsClientDAL (implémentation Access/DAO)
Implements IClientRepository

Private Function IClientRepository_GetById(ByVal clientId As Long) As clsClient
    ' ... même corps qu'auparavant ...
End Function

Private Function IClientRepository_Insert(ByVal client As clsClient) As Long
    ' ...
End Function
' (et ainsi de suite pour Update et EmailExiste)
```

La BLL ne dépend plus alors de la classe concrète mais de l'interface, l'implémentation concrète lui étant fournie au moment de l'instanciation (idéalement par une Factory, section 16.7) :

```vba
' Dans clsClientService : dépendance sur le CONTRAT, pas sur l'implémentation
Private m_repo As IClientRepository

Private Sub Class_Initialize()
    ' La Factory décide quelle DAL fournir (Access aujourd'hui,
    ' SQL Server demain) sans que la BLL ait à le savoir.
    Set m_repo = FactoryDAL.CreerClientRepository()
End Sub
```

Le jour de la migration, il suffit d'écrire une nouvelle classe `clsClientDALSqlServer` qui implémente la même `IClientRepository` en parlant à SQL Server (via tables liées ODBC ou requêtes Pass-Through, sections 23.4 et 11.9), puis de faire pointer la Factory vers elle. La BLL et l'UI ne s'aperçoivent de rien. C'est l'objectif ultime que sert l'architecture en couches : **rendre le changement local et maîtrisé plutôt que global et risqué.**

## Avantages et limites dans le contexte Access

Les bénéfices de cette organisation sont réels. La logique métier est centralisée et donc non dupliquée. Le code devient testable : on peut exercer un `clsClientService` directement depuis la fenêtre d'exécution immédiate ou un harnais de test (sections 19.2 et 19.4), sans ouvrir le moindre formulaire. La maintenance est facilitée, chaque type de changement étant cantonné à une couche. Et la migration vers un autre SGBD devient une opération bornée.

Mais ces avantages ont un coût qu'il faut assumer lucidement. L'architecture en couches **multiplie le nombre de modules** : là où l'anti-modèle tenait dans un seul formulaire, on a désormais une entité, une DAL, une BLL et une UI. Pour une base de quelques formulaires sans logique métier, cette structure est de la sur-ingénierie pure : elle ralentit le développement sans rien apporter. L'instanciation systématique d'objets a par ailleurs un léger coût en performance, généralement négligeable mais réel dans des boucles serrées. Enfin, la nature « RAD » d'Access, conçue pour aller vite avec des formulaires liés, frotte en permanence contre toute abstraction lourde : l'architecture en couches va à contre-courant de la pente naturelle de l'outil, ce qui demande de la discipline pour ne pas régresser vers le « tout dans le formulaire » au premier ajout pressé.

Le bon réflexe n'est donc pas d'appliquer l'architecture en couches partout, mais de la réserver aux applications dont la logique métier est suffisamment riche ou amenée à durer pour que la rigueur soit rentable.

## Pièges courants

Le piège le plus fréquent est la **fuite de la DAL dans l'UI** : un développeur pressé ouvre un `Recordset` directement dans un événement de formulaire « juste pour cette fois ». La frontière est rompue, et la couche d'accès aux données cesse d'être le point de passage unique. La règle doit être absolue : aucun `CurrentDb`, aucun `OpenRecordset`, aucun SQL dans un module de formulaire.

Le piège symétrique est la **BLL qui connaît l'UI** : une classe de service qui référence `Me`, lit `Forms!frmClient!txtNom` ou affiche une `MsgBox`. Dès qu'une classe métier mentionne un contrôle ou une boîte de dialogue, elle perd sa réutilisabilité et sa testabilité. La BLL reçoit des valeurs en paramètres et signale ses refus par `Err.Raise` ; elle n'affiche jamais rien.

Un piège plus subtil est le **modèle anémique** : des entités réduites à de simples sacs de propriétés, toute la logique étant reléguée dans les services. Ce n'est pas faux en soi, mais une entité peut légitimement porter ses propres règles internes (une méthode `clsCommande.CalculerTotal`, par exemple), ce qui allège la BLL et rapproche le code du domaine.

Enfin, le piège inverse de tous les précédents : **sur-architecturer une petite base**. Empiler quatre couches, des interfaces et une Factory pour gérer trois formulaires de saisie est un contre-sens. L'architecture en couches est un investissement ; comme tout investissement, elle ne se justifie que si l'application a la taille et la durée de vie pour le rentabiliser.

## Synthèse

L'architecture en couches consiste à répartir une application Access selon trois rôles stricts : la **présentation (UI)** affiche et capte les actions, la **logique métier (BLL)** porte les règles, l'**accès aux données (DAL)** encapsule la base. Les dépendances ne vont que vers le bas, et une **entité de domaine** circule entre les couches comme contrat partagé. Cette structure centralise la logique, rend le code testable et borne l'impact des évolutions, en particulier d'une éventuelle migration vers SQL Server.

La spécificité d'Access tient à ses formulaires liés, qui court-circuitent naturellement les couches. La réponse n'est pas de les bannir mais de **doser** : formulaires liés pour le CRUD simple, validation routée par la BLL via `BeforeUpdate`, et formulaires indépendants adossés à la BLL pour les opérations complexes. Appliquée avec ce discernement — et réservée aux applications dont la richesse métier la justifie — l'architecture en couches transforme une base Access fragile et difficile à maintenir en une application structurée, durable et prête à évoluer.

---

> 🔗 **Pour aller plus loin :** le patron Repository (section 16.6) détaille l'implémentation de la DAL ; les patrons Singleton et Factory (section 16.7) servent à instancier proprement les couches ; les interfaces (section 16.5) rendent la DAL interchangeable ; la gestion centralisée des erreurs (sections 13.7 à 13.9) accompagne la propagation entre couches ; et le chapitre 23 traite la migration vers SQL Server que cette architecture facilite.

⏭️ [17. Interface utilisateur avancée](/17-interface-utilisateur-avancee/README.md)
