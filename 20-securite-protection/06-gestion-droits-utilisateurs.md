🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 20.6. Gestion applicative des droits utilisateurs (table de rôles)

Puisque le format `.accdb` ne propose plus aucun système de comptes ni de permissions intégré, comme nous l'avons établi en [section 20.1](01-modele-securite-access.md), c'est au développeur de bâtir lui-même la gestion des droits dont son application a besoin. C'est l'objet de cette section : concevoir un système de rôles et de permissions, piloté par des tables et appliqué dans l'interface par code.

Avant tout, il faut situer correctement ce que cette couche protège — et ce qu'elle ne protège pas. Il s'agit d'une sécurité fonctionnelle : elle contrôle ce qu'un utilisateur a le droit de faire à l'intérieur de l'application (créer, modifier, supprimer, accéder à tel module). Elle ne contrôle ni l'accès au fichier sur le disque, qui relève du chiffrement par mot de passe (voir [section 20.2](02-protection-mot-de-passe.md)), ni la lecture du code, qui relève du `.accde` (voir [section 20.3](03-compilation-accde.md)). S'exécutant à l'intérieur de l'application, elle n'a de valeur que si l'utilisateur ne peut pas la contourner. Un utilisateur qui ouvrirait directement le `.accdb`, accéderait aux tables ou modifierait le code rendrait tout ce dispositif inopérant. La gestion applicative des droits doit donc impérativement être combinée avec le `.accde`, le chiffrement du back-end et le verrouillage de l'interface au démarrage (voir [section 17.7](../17-interface-utilisateur-avancee/07-options-demarrage.md)). C'est un outil de gouvernance et de prévention des usages indus ou accidentels par des utilisateurs normaux, et non une barrière infranchissable contre un utilisateur technique disposant d'un accès au fichier.

## Le modèle de données : tables de rôles et de droits

Le modèle retenu est un contrôle d'accès fondé sur les rôles (en anglais *Role-Based Access Control*). Plutôt que d'attribuer des permissions individuellement à chaque utilisateur, on définit des rôles, on rattache des droits à ces rôles, et l'on assigne un rôle à chaque utilisateur. Ce modèle repose sur quatre tables.

Une table des rôles décrit les profils (par exemple Administrateur, Gestionnaire, Consultation). Une table des droits énumère les permissions élémentaires, chacune identifiée par un code textuel parlant comme `CLIENTS_MODIFIER` ou `FACTURES_SUPPRIMER` ; ces codes serviront de point d'ancrage dans le code VBA. Une table de jonction relie rôles et droits, traduisant le fait qu'un rôle possède un ensemble de droits. Enfin, une table des utilisateurs recense les comptes et rattache chacun à un rôle.

Les instructions DDL suivantes (dialecte Jet, voir [section 11.4](../11-sql-access-vba/04-requetes-ddl.md)) illustrent cette structure ; ces tables peuvent tout aussi bien être créées dans le concepteur de tables.

```sql
CREATE TABLE tblRoles (
    RoleID      AUTOINCREMENT PRIMARY KEY,
    NomRole     TEXT(50)  NOT NULL,
    Description TEXT(255)
);

CREATE TABLE tblDroits (
    DroitID   AUTOINCREMENT PRIMARY KEY,
    CodeDroit TEXT(50)  NOT NULL UNIQUE,
    Libelle   TEXT(255)
);

CREATE TABLE tblUtilisateurs (
    UtilisateurID      AUTOINCREMENT PRIMARY KEY,
    IdentifiantWindows TEXT(100) NOT NULL,
    NomComplet         TEXT(150),
    Actif              BIT       NOT NULL,
    RoleID             LONG REFERENCES tblRoles (RoleID)
);

CREATE TABLE tblRolesDroits (
    RoleID  LONG NOT NULL REFERENCES tblRoles  (RoleID),
    DroitID LONG NOT NULL REFERENCES tblDroits (DroitID),
    CONSTRAINT pk_RolesDroits PRIMARY KEY (RoleID, DroitID)
);
```

Ce modèle attribue un rôle unique par utilisateur, ce qui suffit dans la plupart des cas et reste simple à administrer. Si une application exige qu'un utilisateur cumule plusieurs rôles, on remplace le champ `RoleID` de la table des utilisateurs par une seconde table de jonction reliant utilisateurs et rôles, sur le même principe que `tblRolesDroits`.

Ces tables ont naturellement vocation à résider dans le back-end partagé, afin que tous les utilisateurs partagent le même référentiel de rôles. Cela rend d'autant plus essentielle la protection du back-end : sans chiffrement, un utilisateur pourrait ouvrir la base de données et s'octroyer lui-même le rôle d'administrateur.

## Identifier l'utilisateur courant

Pour appliquer des droits, encore faut-il savoir qui est connecté. Deux approches existent.

La plus simple et la plus transparente consiste à s'appuyer sur l'identifiant de session Windows. L'utilisateur n'a aucun mot de passe supplémentaire à retenir, et l'on évite d'avoir à stocker et vérifier des mots de passe applicatifs. On récupère cet identifiant via la variable d'environnement `USERNAME`, ou, de façon plus robuste, par l'API Windows `GetUserName` présentée en [section 22.2](../22-api-windows-integration-office/02-api-courantes.md). Il suffit ensuite de retrouver la ligne correspondante dans la table des utilisateurs.

La seconde approche repose sur un formulaire d'identification applicatif, avec un identifiant et un mot de passe propres à l'application. Elle affranchit la gestion des droits des comptes Windows, ce qui peut être utile sur des postes partagés, mais elle impose de stocker les mots de passe de façon sécurisée, c'est-à-dire jamais en clair : on en conserve une empreinte (hachage), comme exposé en [section 20.8](08-chiffrement-donnees-sensibles.md). Pour la majorité des applications, l'identification par session Windows est recommandée.

## Charger les droits au démarrage

L'idée directrice est de déterminer l'utilisateur et de charger l'ensemble de ses droits une seule fois, au démarrage de l'application, dans une structure en mémoire offrant une vérification instantanée. Un dictionnaire (`Scripting.Dictionary`, voir [section 3.7](../03-rappels-fondamentaux/07-collections-dictionary.md)) est idéal pour cet usage, car sa méthode `Exists` répond immédiatement.

Le module ci-dessous, que l'on nommera `modSecurite`, regroupe cette logique. La fonction `InitialiserSession` est appelée une fois au lancement (depuis une macro `AutoExec` ou un formulaire de démarrage).

```vba
' === Module : modSecurite ===
Option Compare Database
Option Explicit

' État de la session, en mémoire pour toute la durée d'exécution
Private mdicDroits As Object          ' Scripting.Dictionary (liaison tardive)
Private mlngUtilisateurID As Long
Private mstrNomComplet As String
Private mblnSessionOuverte As Boolean

Public Function InitialiserSession() As Boolean
    ' Identifie l'utilisateur Windows actif, charge son rôle et ses droits.
    On Error GoTo Echec

    Dim strLogin As String
    strLogin = Environ$("USERNAME")    ' voir 22.2 pour l'API GetUserName (plus robuste)

    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim strSQL As String
    Set db = CurrentDb

    ' 1) Retrouver l'utilisateur actif correspondant à l'identifiant Windows
    strSQL = "SELECT UtilisateurID, NomComplet, RoleID " & _
             "FROM tblUtilisateurs " & _
             "WHERE Actif = True AND IdentifiantWindows = '" & _
             Replace$(strLogin, "'", "''") & "'"
    Set rs = db.OpenRecordset(strSQL, dbOpenSnapshot)

    If rs.EOF Then
        mblnSessionOuverte = False      ' utilisateur inconnu ou désactivé
        InitialiserSession = False
        GoTo Nettoyage
    End If

    mlngUtilisateurID = rs!UtilisateurID
    mstrNomComplet = Nz(rs!NomComplet, strLogin)

    Dim lngRoleID As Long
    lngRoleID = Nz(rs!RoleID, 0)
    rs.Close

    ' 2) Charger les codes de droits du rôle dans un dictionnaire
    Set mdicDroits = CreateObject("Scripting.Dictionary")
    mdicDroits.CompareMode = vbTextCompare   ' comparaison insensible à la casse

    strSQL = "SELECT d.CodeDroit " & _
             "FROM tblDroits AS d " & _
             "INNER JOIN tblRolesDroits AS rd ON d.DroitID = rd.DroitID " & _
             "WHERE rd.RoleID = " & lngRoleID
    Set rs = db.OpenRecordset(strSQL, dbOpenSnapshot)
    Do While Not rs.EOF
        If Not mdicDroits.Exists(rs!CodeDroit & "") Then
            mdicDroits.Add rs!CodeDroit & "", True
        End If
        rs.MoveNext
    Loop

    mblnSessionOuverte = True
    InitialiserSession = True

Nettoyage:
    On Error Resume Next
    rs.Close
    Set rs = Nothing
    Set db = Nothing
    Exit Function

Echec:
    mblnSessionOuverte = False
    InitialiserSession = False
    Resume Nettoyage
End Function
```

On complète ce module par quelques propriétés d'accès en lecture seule à l'identité de l'utilisateur courant, utiles pour l'affichage et pour la journalisation d'audit (voir [section 20.7](07-audit-acces-modifications.md)).

```vba
Public Property Get UtilisateurID() As Long
    UtilisateurID = mlngUtilisateurID
End Property

Public Property Get NomComplet() As String
    NomComplet = mstrNomComplet
End Property

Public Property Get SessionOuverte() As Boolean
    SessionOuverte = mblnSessionOuverte
End Property
```

## La fonction centrale `ADroit`

Tout le système repose sur une unique fonction de vérification, simple et rapide, qui sera invoquée partout dans l'application. Elle répond à la question : l'utilisateur courant possède-t-il tel droit ?

```vba
Public Function ADroit(ByVal codeDroit As String) As Boolean
    ' Cœur du système : l'utilisateur courant possède-t-il ce droit ?
    If Not mblnSessionOuverte Then
        ADroit = False
    ElseIf mdicDroits Is Nothing Then
        ADroit = False
    Else
        ADroit = mdicDroits.Exists(codeDroit)
    End If
End Function
```

Le choix par défaut est volontairement restrictif : en l'absence de session ou de dictionnaire chargé, la fonction renvoie `False`. Aucun droit n'est accordé tant que l'identité n'a pas été établie, ce qui est le comportement prudent attendu d'un système de sécurité.

## Appliquer les droits dans l'interface

Munie de `ADroit`, l'application adapte son interface au profil connecté. La forme la plus directe consiste à activer ou désactiver des contrôles dans l'événement de chargement d'un formulaire.

```vba
Private Sub Form_Load()
    Me.cmdNouveau.Enabled = ADroit("CLIENTS_CREER")
    Me.cmdModifier.Enabled = ADroit("CLIENTS_MODIFIER")
    Me.cmdSupprimer.Enabled = ADroit("CLIENTS_SUPPRIMER")
End Sub
```

Lorsque les contrôles à régir sont nombreux, cette écriture devient fastidieuse. Une approche plus élégante et plus maintenable consiste à déclarer le droit requis directement dans la propriété `Tag` du contrôle, selon une convention (par exemple `DROIT=CLIENTS_SUPPRIMER`), puis à parcourir les contrôles du formulaire pour appliquer la règle automatiquement.

```vba
Public Sub AppliquerDroits(ByRef frm As Form)
    ' Masque tout contrôle dont le Tag réclame un droit que l'utilisateur n'a pas.
    Dim ctl As Control
    Dim strCode As String
    For Each ctl In frm.Controls
        strCode = ExtraireDroitDuTag(ctl.Tag)
        If Len(strCode) > 0 Then
            On Error Resume Next       ' tous les contrôles n'exposent pas Visible
            ctl.Visible = ADroit(strCode)
            On Error GoTo 0
        End If
    Next ctl
End Sub

Private Function ExtraireDroitDuTag(ByVal strTag As String) As String
    ' Extrait le code après "DROIT=" dans la propriété Tag
    Const PREFIXE As String = "DROIT="
    Dim p As Long
    p = InStr(1, strTag, PREFIXE, vbTextCompare)
    If p = 0 Then Exit Function

    Dim reste As String
    reste = Mid$(strTag, p + Len(PREFIXE))
    Dim sep As Long
    sep = InStr(reste, ";")            ' le code s'arrête à un éventuel ";"
    If sep > 0 Then reste = Left$(reste, sep - 1)
    ExtraireDroitDuTag = Trim$(reste)
End Function
```

Il suffit alors d'appeler `AppliquerDroits Me` au chargement de chaque formulaire concerné. La logique de droits reste centralisée, et l'ajout d'un nouveau contrôle protégé ne demande qu'à renseigner sa propriété `Tag`.

## Défense en profondeur : vérifier à plusieurs niveaux

Une erreur fréquente consiste à se contenter de masquer ou de désactiver un bouton. Or rien ne garantit que l'action correspondante ne sera pas déclenchée autrement. Il faut donc, en plus de l'adaptation visuelle, revérifier le droit à l'entrée même de l'action.

```vba
Private Sub cmdSupprimer_Click()
    ' Ne jamais se fier au seul masquage : on contrôle aussi à l'exécution.
    If Not ADroit("CLIENTS_SUPPRIMER") Then
        MsgBox "Vous n'avez pas le droit de supprimer un client.", vbExclamation
        Exit Sub
    End If
    ' ... suppression effective ...
End Sub
```

Ce double contrôle — adapter l'interface, puis vérifier avant d'agir — est la règle d'or de la sécurité applicative. La centralisation de toute la logique dans `modSecurite` garantit par ailleurs qu'une éventuelle évolution des règles se fait en un seul endroit.

## Administration des utilisateurs et des rôles

Pour être exploitable au quotidien, ce dispositif s'accompagne de formulaires d'administration permettant de gérer les utilisateurs, de définir les rôles et d'attribuer les droits à chaque rôle. Ces écrans, réservés au seul profil administrateur (eux-mêmes protégés par un droit du type `ADMIN_SECURITE`), ne présentent pas de difficulté particulière : ce sont des formulaires de saisie classiques opérant sur les quatre tables décrites plus haut. L'essentiel est qu'ils soient inaccessibles aux utilisateurs ordinaires.

## Mise en cache et rafraîchissement

Charger les droits une seule fois au démarrage offre d'excellentes performances, puisque chaque appel à `ADroit` se résout en mémoire sans accès à la base. La contrepartie est qu'une modification des droits d'un utilisateur déjà connecté ne prend effet qu'à sa prochaine ouverture de session. Si ce délai pose problème, on peut prévoir une procédure de rechargement réappelant `InitialiserSession`. Cette stratégie de cache rejoint les principes vus en [section 18.4](../18-optimisation-performance/04-cache-variables-module.md), et l'on peut, en variante, stocker l'identité courante dans des `TempVars` (voir [section 15.7](../15-multi-utilisateurs/07-tempvars.md)) pour la partager aisément entre formulaires.

## Points de vigilance

Le point décisif a été énoncé en ouverture et mérite d'être répété : cette sécurité s'exécute dans l'application et ne vaut que par les protections qui l'entourent. Sans `.accde`, le code des vérifications peut être altéré ; sans chiffrement du back-end, les tables de rôles peuvent être modifiées directement ; sans verrouillage du démarrage, l'utilisateur peut accéder aux données en contournant l'interface. Ces couches sont indissociables.

Veillez en outre à conserver le comportement restrictif par défaut — aucun droit sans identité établie — et à toujours prévoir le cas de l'utilisateur inconnu, qui ne doit obtenir aucun accès. Enfin, la gestion des droits gagne à être complétée par un audit des actions sensibles, traité dans la section suivante, afin de savoir non seulement qui avait le droit d'agir, mais qui a effectivement agi.

## Ce qu'il faut retenir

En l'absence de sécurité native, la gestion des droits dans une application Access se construit autour d'un modèle de rôles et de permissions porté par quatre tables, d'une identification de l'utilisateur (de préférence via la session Windows), d'un chargement des droits en mémoire au démarrage, et d'une fonction centrale `ADroit` invoquée partout. L'interface s'adapte au profil, et chaque action sensible revérifie le droit avant de s'exécuter. Cette couche de sécurité fonctionnelle, précieuse pour structurer et gouverner les accès des utilisateurs normaux, n'a toutefois de portée réelle que combinée au `.accde`, au chiffrement du back-end et au verrouillage de l'interface étudiés dans les autres sections de ce chapitre.

---


⏭️ [20.7. Audit des accès et des modifications par code](/20-securite-protection/07-audit-acces-modifications.md)
