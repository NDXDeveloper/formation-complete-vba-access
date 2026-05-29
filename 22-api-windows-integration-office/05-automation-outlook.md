🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 22.5. Automation avec Outlook — envoi d'emails et gestion des contacts

L'automation d'Outlook répond à deux besoins très courants dans une application Access : **envoyer des courriels** — confirmations, relances, états joints en pièce jointe — et **gérer des contacts** entre Access et le carnet d'adresses. C'est souvent l'aboutissement de chaînes de traitement entamées ailleurs : un état exporté en PDF (section 7.8) ou un document généré sous Word (section 22.4) que l'on expédie ensuite par message.

Cette section s'appuie sur les fondations posées à la section 22.3, mais Outlook présente **deux spécificités majeures** qui le distinguent d'Excel et de Word, et qu'il faut maîtriser avant tout : son caractère mono-instance, et son garde-barrière de sécurité.

## Les fondations, et deux spécificités Outlook

Les principes de la section 22.3 demeurent : liaison tardive privilégiée, réglages de l'application, et libération des objets. Le modèle objet d'Outlook s'articule autour de l'objet `Application`, de la méthode `CreateItem` (qui crée un message, un contact, un rendez-vous…), et de l'espace de noms obtenu par `GetNamespace("MAPI")`, porte d'entrée vers les dossiers via `GetDefaultFolder`. Deux particularités appellent toutefois une vigilance propre à Outlook.

### Outlook est mono-instance : ne pas appeler Quit

Contrairement à Excel et Word, Outlook est une application **mono-instance** : `CreateObject("Outlook.Application")` ne crée pas une nouvelle instance si Outlook tourne déjà, mais **renvoie l'instance en cours** — celle de l'utilisateur. La conséquence est directe : **il ne faut pas appeler `Quit` sur Outlook**, sous peine de fermer la session Outlook de l'utilisateur en pleine utilisation.

La routine de nettoyage se limite donc, pour Outlook, à **libérer les variables objets** (`Set ... = Nothing`), sans jamais quitter l'application. C'est une différence essentielle avec le squelette d'automation d'Excel ou de Word.

### Le garde-barrière de sécurité

La seconde spécificité est le **modèle de sécurité d'Outlook** (Object Model Guard), conçu pour empêcher un programme malveillant d'envoyer des messages en masse ou de piller le carnet d'adresses. L'accès programmatique à l'envoi ou aux adresses peut donc déclencher des **invites de sécurité** demandant l'autorisation de l'utilisateur.

Le comportement de ce garde-barrière **varie selon l'environnement** : présence d'un antivirus à jour et reconnu par Windows, paramètres de stratégie de groupe, version d'Outlook. Dans un parc d'entreprise bien administré, les invites peuvent être supprimées par configuration ; mais sur un poste quelconque, on ne peut jamais présumer leur absence. Plusieurs approches permettent de composer avec cette contrainte :

- Afficher le message par `.Display` plutôt que de l'envoyer par `.Send` : c'est alors **l'utilisateur** qui clique sur Envoyer, ce qui évite l'invite liée à l'envoi programmatique.
- Configurer l'accès programmatique par **stratégie de groupe**, dans les environnements gérés.
- Recourir à une bibliothèque tierce spécialisée qui contourne le garde-barrière, au prix d'une dépendance externe.
- **Se passer d'Outlook** et envoyer directement par SMTP, comme exposé plus bas — souvent la meilleure réponse pour un envoi automatisé non surveillé.

## Envoyer un email

Un message se crée par `CreateItem` avec le type `olMailItem` (valeur 0 à définir en liaison tardive). On en renseigne ensuite les propriétés, on ajoute d'éventuelles pièces jointes, puis on l'affiche ou on l'envoie :

```vba
Public Sub EnvoyerEmail(ByVal destinataire As String, _
                        ByVal sujet As String, _
                        ByVal corpsHTML As String, _
                        Optional ByVal pieceJointe As String = "")
    Const olMailItem As Long = 0

    Dim olApp As Object, mail As Object

    On Error GoTo Nettoyage

    Set olApp = CreateObject("Outlook.Application")
    Set mail = olApp.CreateItem(olMailItem)

    With mail
        .To = destinataire
        .Subject = sujet
        .HTMLBody = corpsHTML
        If Len(pieceJointe) > 0 Then .Attachments.Add pieceJointe
        .Display          ' afficher pour relecture ; .Send pour envoyer directement
    End With

Nettoyage:
    On Error Resume Next
    ' Ne PAS appeler olApp.Quit : Outlook est mono-instance
    Set mail = Nothing
    Set olApp = Nothing
End Sub
```

### Corps texte ou HTML

Le corps du message se définit soit en texte brut via la propriété `.Body`, soit en **HTML** via `.HTMLBody`, qui autorise une mise en forme riche (paragraphes, gras, liens, tableaux). Renseigner `.HTMLBody` bascule automatiquement le message au format HTML.

### Pièces jointes

La méthode `.Attachments.Add` joint un fichier au message. C'est ici que se referme la chaîne de traitement évoquée en introduction : on génère un état en PDF (section 7.8) ou un document Word (section 22.4), puis on le joint au message pour l'expédier.

### Envoyer ou afficher : .Send et .Display

Le choix entre `.Send` et `.Display` structure tout l'usage :

- **`.Send`** expédie le message immédiatement, sans intervention. C'est l'option des traitements automatisés — mais c'est aussi celle qui sollicite le plus le garde-barrière de sécurité.
- **`.Display`** ouvre le message à l'écran pour que l'utilisateur le relise, le complète éventuellement, puis l'envoie lui-même. C'est l'option des envois supervisés, et elle évite l'invite liée à l'envoi programmatique.

## Envoi en nombre personnalisé

Un usage classique consiste à envoyer un message **personnalisé par destinataire**, à partir d'une requête : on parcourt le jeu d'enregistrements (chapitre 9) et l'on crée un message par client. Chaque message est libéré dans la boucle pour éviter toute accumulation :

```vba
Public Sub EnvoyerCampagne(ByVal sql As String)
    Const olMailItem As Long = 0
    Dim olApp As Object, mail As Object
    Dim db As DAO.Database, rs As DAO.Recordset

    On Error GoTo Nettoyage
    Set db = CurrentDb
    Set rs = db.OpenRecordset(sql, dbOpenSnapshot)
    Set olApp = CreateObject("Outlook.Application")

    Do Until rs.EOF
        Set mail = olApp.CreateItem(olMailItem)
        With mail
            .To = Nz(rs!Email, "")
            .Subject = "Information sur votre compte"
            .HTMLBody = "<p>Bonjour " & Nz(rs!Nom, "") & ",</p>" & _
                        "<p>Votre solde est de " & _
                        Format$(Nz(rs!Solde, 0), "#,##0.00 €") & ".</p>"
            .Send
        End With
        Set mail = Nothing       ' libérer à chaque itération
        rs.MoveNext
    Loop

Nettoyage:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set mail = Nothing
    Set olApp = Nothing          ' pas de Quit
    Set rs = Nothing
    Set db = Nothing
End Sub
```

C'est précisément ce scénario d'envoi programmatique en nombre que le garde-barrière de sécurité cible : selon la configuration du poste, il peut déclencher des invites répétées. Pour un envoi de masse réellement non surveillé, l'approche SMTP ci-dessous est souvent préférable. Cet usage rejoint la section 7.10, consacrée aux états paramétrés et à leur envoi automatisé par courriel.

## L'alternative sans Outlook : l'envoi par SMTP

Lorsqu'Outlook ne convient pas — absent du poste, envoi non surveillé, ou garde-barrière à éviter —, on peut envoyer directement par **SMTP**, sans aucune dépendance à Outlook. Le composant **CDO**, présent en standard sous Windows, le permet : on configure le serveur SMTP, puis on envoie le message.

```vba
Public Sub EnvoyerParSMTP(ByVal destinataire As String, _
                          ByVal sujet As String, _
                          ByVal corpsHTML As String)
    Const cdoSendUsingPort As Long = 2
    Const SCHEMA As String = "http://schemas.microsoft.com/cdo/configuration/"

    Dim msg As Object, conf As Object, champs As Object

    On Error GoTo Gestion

    Set msg = CreateObject("CDO.Message")
    Set conf = CreateObject("CDO.Configuration")
    Set champs = conf.Fields

    With champs
        .Item(SCHEMA & "sendusing") = cdoSendUsingPort
        .Item(SCHEMA & "smtpserver") = "smtp.exemple.fr"
        .Item(SCHEMA & "smtpserverport") = 25
        .Update
    End With

    With msg
        Set .Configuration = conf
        .From = "expediteur@exemple.fr"
        .To = destinataire
        .Subject = sujet
        .HTMLBody = corpsHTML
        .Send
    End With

Gestion:
    Set champs = Nothing
    Set conf = Nothing
    Set msg = Nothing
End Sub
```

Pour un serveur exigeant une authentification ou une connexion chiffrée, on ajoute les champs correspondants (authentification, nom d'utilisateur, mot de passe, chiffrement). Une réserve d'honnêteté s'impose toutefois : CDO est une technologie ancienne, qui dialogue bien avec un **relais SMTP interne** mais s'accommode mal des fournisseurs modernes imposant des protocoles de chiffrement récents. Dans ce dernier cas, l'envoi via une **API web** (par exemple un service d'emailing appelé en REST, section 22.7) constitue une alternative plus robuste.

## Gérer les contacts

### Créer un contact

Un contact se crée par `CreateItem` avec le type `olContactItem` (valeur 2). On renseigne ses propriétés, puis `.Save` l'enregistre dans le dossier Contacts par défaut :

```vba
Public Sub CreerContact(ByVal nomComplet As String, ByVal email As String, _
                        ByVal societe As String)
    Const olContactItem As Long = 2
    Dim olApp As Object, contact As Object

    On Error GoTo Nettoyage
    Set olApp = CreateObject("Outlook.Application")
    Set contact = olApp.CreateItem(olContactItem)

    With contact
        .FullName = nomComplet
        .Email1Address = email
        .CompanyName = societe
        .Save                ' enregistre dans le dossier Contacts par défaut
    End With

Nettoyage:
    On Error Resume Next
    Set contact = Nothing
    Set olApp = Nothing
End Sub
```

### Lire les contacts

Pour importer les contacts d'Outlook dans Access, on accède au **dossier Contacts** via l'espace de noms MAPI (`GetDefaultFolder` avec le code 10), puis on parcourt sa collection `Items` :

```vba
Public Sub ImporterContacts()
    Const olFolderContacts As Long = 10
    Const olContact As Long = 40    ' valeur de la propriété Class pour un contact

    Dim olApp As Object, ns As Object, dossier As Object, item As Object
    Dim db As DAO.Database, rs As DAO.Recordset

    On Error GoTo Nettoyage
    Set olApp = CreateObject("Outlook.Application")
    Set ns = olApp.GetNamespace("MAPI")
    Set dossier = ns.GetDefaultFolder(olFolderContacts)

    Set db = CurrentDb
    Set rs = db.OpenRecordset("tblContacts", dbOpenDynaset, dbAppendOnly)

    For Each item In dossier.Items
        If item.Class = olContact Then       ' filtrer les seuls contacts
            rs.AddNew
            rs!Nom = item.FullName
            rs!Email = item.Email1Address
            rs!Societe = item.CompanyName
            rs.Update
        End If
    Next item

Nettoyage:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set item = Nothing
    Set dossier = Nothing
    Set ns = Nothing
    Set olApp = Nothing
    Set rs = Nothing
    Set db = Nothing
End Sub
```

Un détail mérite attention : la propriété `Class` de l'élément relève d'une **énumération différente** de celle utilisée par `CreateItem`. Un contact se crée avec le type `olContactItem` (valeur 2), mais sa propriété `Class` vaut `olContact` (valeur 40). Filtrer sur cette valeur écarte au passage les listes de distribution, que le dossier Contacts peut également renfermer.

## Articulation avec le reste du chapitre et de la formation

Cette section prolonge l'automation Office et referme plusieurs chaînes de traitement :

- Les **fondations de l'automation** sont établies à la section 22.3, et le choix de liaison à la section 2.6.
- L'**envoi d'un document généré** s'articule avec Word (section 22.4) et l'export d'états en PDF (section 7.8).
- L'**envoi automatisé d'états paramétrés** par courriel relève de la section 7.10.
- L'**alternative par API web** au SMTP renvoie à la section 22.7.
- Les **jeux d'enregistrements** alimentant les envois et l'import de contacts relèvent du chapitre 9, la **gestion d'erreurs** du chapitre 13, et les **implications de sécurité** de la section 20.5.

## Points de vigilance

Plusieurs erreurs reviennent fréquemment dans l'automation d'Outlook :

- **Appeler `Quit` sur Outlook**, et fermer la session de l'utilisateur, Outlook étant mono-instance.
- **Présumer l'absence d'invites de sécurité**, alors que le garde-barrière dépend de la configuration du poste.
- **Recourir à `.Send` pour un envoi de masse** sans tenir compte du garde-barrière, là où le SMTP serait plus adapté.
- **Compter sur CDO avec un fournisseur moderne** exigeant un chiffrement récent, que cette technologie ancienne gère mal.
- **Confondre les énumérations** `olContactItem` (création, valeur 2) et `olContact` (propriété `Class`, valeur 40) lors du filtrage des éléments.
- **Ne pas libérer les messages dans une boucle d'envoi**, laissant les objets s'accumuler.

⏭️ [22.6. Automation avec PowerPoint — génération de présentations](/22-api-windows-integration-office/06-automation-powerpoint.md)
