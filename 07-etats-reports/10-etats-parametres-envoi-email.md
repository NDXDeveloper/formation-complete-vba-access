🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 7.10. États paramétrés et envoi par email automatisé

Cette section clôt le chapitre en réunissant ses acquis pour automatiser la **diffusion** des états. L'idée : à partir d'un **seul état paramétré**, produire un document filtré pour chaque destinataire (chapitre 7.4), l'exporter en PDF (chapitre 7.8), puis l'**envoyer par e-mail** — éventuellement **en lot**, un envoi par client. C'est le scénario classique de l'envoi automatisé de factures, relevés ou bulletins.

Deux briques nouvelles s'ajoutent ici aux mécanismes déjà vus : l'envoi par `DoCmd.SendObject` et l'envoi par **automation Outlook**.

---

## 7.10.1. Le principe : un état, des sorties multiples

Un **état paramétré** est un état unique dont on fait varier le périmètre par code. Plutôt que de concevoir un état par client, on conçoit un état générique et on l'alimente, à chaque production, d'un **filtre** (le client, la période) et éventuellement d'un **contexte d'affichage** (titre, libellé de période).

L'enchaînement complet de la diffusion automatisée est :

```
Paramétrer (filtre)  →  Exporter en PDF  →  Joindre à un e-mail  →  Envoyer
   (chapitre 7.4)         (chapitre 7.8)              (cette section)
```

Répété en boucle sur une liste de destinataires, il réalise l'**envoi en lot** (section 7.10.8).

---

## 7.10.2. Paramétrer un état (rappel et consolidation)

Les mécanismes de paramétrage ont été détaillés plus tôt ; on les mobilise ici :

- **Filtrer** le contenu : argument `WhereCondition` de `DoCmd.OpenReport` (chapitre 7.4) ;
- **Transmettre un contexte d'affichage** (titre, période) : argument `OpenArgs` (chapitre 6.6), lu dans l'événement `Open` ;
- **Adapter le tri** : collection `GroupLevel` (chapitre 7.5).

```vba
' État filtré + contexte transmis via OpenArgs
DoCmd.OpenReport "E_Facture", acViewPreview, , _
                 "ClientID=" & lngID, acHidden, "Facture du " & Format(Date, "dd/mm/yyyy")
```

L'essentiel de cette section porte sur ce qui vient **après** : l'envoi.

---

## 7.10.3. Deux voies pour l'envoi : SendObject ou Outlook

| Voie | Avantages | Limites |
|---|---|---|
| `DoCmd.SendObject` | Intégré, simple, sans dépendance | Peu d'options, une seule pièce jointe, avertissements de sécurité |
| Automation **Outlook** | Contrôle total (CC, CCi, HTML, pièces multiples) | Nécessite Outlook installé |

`SendObject` convient aux envois simples ; l'**automation Outlook** s'impose dès qu'on veut joindre un PDF pré-exporté, soigner le corps en HTML, ou gérer plusieurs destinataires et pièces jointes.

---

## 7.10.4. DoCmd.SendObject

`DoCmd.SendObject` produit l'objet (ici l'état) et l'attache à un message envoyé via le client de messagerie par défaut :

```vba
DoCmd.SendObject ObjectType, [ObjectName], [OutputFormat], [To], [Cc], [Bcc], _
                 [Subject], [MessageText], [EditMessage], [TemplateFile]
```

| Argument | Rôle |
|---|---|
| `ObjectType` | `acSendReport` (ou `acSendNoObject` pour un message sans pièce jointe) |
| `OutputFormat` | Format de la pièce jointe : `acFormatPDF`… |
| `To` / `Cc` / `Bcc` | Destinataires (séparés par des points-virgules) |
| `Subject` / `MessageText` | Objet et corps du message |
| `EditMessage` | `True` : ouvre le message pour relecture ; `False` : envoie directement |

```vba
DoCmd.SendObject acSendReport, "E_Facture", acFormatPDF, _
    "client@exemple.com", , , "Votre facture", _
    "Bonjour, veuillez trouver votre facture en pièce jointe.", False
```

> **Même limite qu'`OutputTo`** : `SendObject` n'accepte **pas** de `WhereCondition`. Pour envoyer un état **filtré**, on l'ouvre d'abord filtré et caché, puis on l'envoie, puis on le referme — exactement comme à la section 7.8.4 :
>
> ```vba
> DoCmd.OpenReport "E_Facture", acViewPreview, , "ClientID=" & lngID, acHidden
> DoCmd.SendObject acSendReport, "E_Facture", acFormatPDF, strEmail, , , _
>                  strSujet, strCorps, False
> DoCmd.Close acReport, "E_Facture"
> ```

---

## 7.10.5. L'avertissement de sécurité et ses implications

`SendObject` s'appuie sur le client de messagerie par défaut (mécanisme MAPI). Selon la configuration, **Outlook peut afficher un avertissement de sécurité** (« Un programme tente d'envoyer un message… ») exigeant une confirmation manuelle — ce qui ruine l'automatisation en lot.

C'est l'une des raisons de préférer l'**automation Outlook** pour les envois automatisés : selon l'environnement et les paramètres de sécurité, elle permet souvent un envoi sans interruption. Ces questions de sécurité et de configuration sont approfondies au chapitre 22.5.

---

## 7.10.6. Envoi via automation Outlook

L'automation Outlook offre le contrôle complet. On crée un objet `Outlook.Application`, un élément de message, on le renseigne et on l'envoie. La **liaison tardive** (`CreateObject`) évite d'imposer une référence à la bibliothèque Outlook (chapitre 2.6) :

```vba
Sub EnvoyerEmailOutlook(ByVal strTo As String, ByVal strSujet As String, _
                        ByVal strCorps As String, ByVal strPJ As String)
    Dim olApp As Object, olMail As Object

    Set olApp = CreateObject("Outlook.Application")
    Set olMail = olApp.CreateItem(0)        ' 0 = élément de message

    With olMail
        .To = strTo
        .Subject = strSujet
        .Body = strCorps                     ' ou .HTMLBody pour un corps HTML
        If Len(strPJ) > 0 And Dir(strPJ) <> "" Then .Attachments.Add strPJ
        .Send                                ' ou .Display pour relire avant envoi
    End With

    Set olMail = Nothing
    Set olApp = Nothing
End Sub
```

Ici, on joint un **PDF déjà exporté** (chapitre 7.8) plutôt que de laisser la messagerie produire la pièce jointe : c'est plus souple et plus fiable. L'automation Outlook est traitée en détail au chapitre 22.5.

---

## 7.10.7. Le pattern complet : état paramétré → PDF → email

On combine les briques dans une procédure produisant et envoyant la facture d'**un** client, dotée de sa propre gestion d'erreurs. Elle ignore les clients sans données (approche proactive du chapitre 7.9) et garantit la fermeture de l'état caché :

```vba
Function EnvoyerFactureClient(ByVal lngID As Long, ByVal strEmail As String) As Boolean
    Dim strChemin As String
    On Error GoTo Gestion
    EnvoyerFactureClient = False

    ' Ignorer si aucune donnée à facturer (proactif, chapitre 7.9)
    If DCount("*", "T_Factures", "ClientID=" & lngID) = 0 Then Exit Function

    ' Exporter le PDF filtré (pattern du chapitre 7.8)
    strChemin = Environ("TEMP") & "\Facture_" & lngID & ".pdf"
    DoCmd.OpenReport "E_Facture", acViewPreview, , "ClientID=" & lngID, acHidden
    DoCmd.OutputTo acOutputReport, "E_Facture", acFormatPDF, strChemin
    DoCmd.Close acReport, "E_Facture"

    ' Envoyer via Outlook avec le PDF en pièce jointe
    EnvoyerEmailOutlook strEmail, "Votre facture", _
                        "Bonjour, votre facture est en pièce jointe.", strChemin

    ' Nettoyer le fichier temporaire
    If Dir(strChemin) <> "" Then Kill strChemin

    EnvoyerFactureClient = True
    Exit Function

Gestion:
    ' Refermer l'état s'il est resté ouvert
    If CurrentProject.AllReports("E_Facture").IsLoaded Then
        DoCmd.Close acReport, "E_Facture"
    End If
    EnvoyerFactureClient = False
End Function
```

---

## 7.10.8. Envoi en lot (boucle sur les destinataires)

L'envoi de masse consiste à parcourir un `Recordset` (chapitre 9) de destinataires et à appeler la procédure unitaire pour chacun. Chaque échec est compté sans interrompre le lot :

```vba
Sub EnvoyerToutesFactures()
    Dim rs As DAO.Recordset
    Dim lngOK As Long, lngKo As Long

    Set rs = CurrentDb.OpenRecordset( _
        "SELECT ClientID, Email FROM T_Clients " & _
        "WHERE Email Is Not Null AND AFacturer = True", dbOpenSnapshot)

    Do While Not rs.EOF
        If EnvoyerFactureClient(rs!ClientID, rs!Email) Then
            lngOK = lngOK + 1
        Else
            lngKo = lngKo + 1
        End If
        rs.MoveNext
    Loop

    rs.Close
    Set rs = Nothing

    MsgBox lngOK & " envoi(s) réussi(s), " & lngKo & " échec(s) ou ignoré(s).", _
           vbInformation, "Envoi des factures"
End Sub
```

Points de conception pour un envoi en lot fiable :

- **gérer les erreurs par itération** (procédure unitaire à gestion propre) pour qu'un incident n'arrête pas tout le lot ;
- **filtrer les destinataires en amont** (adresse non nulle, données présentes) ;
- **vérifier la validité** de l'adresse avant l'envoi ;
- **indiquer la progression** sur de gros volumes (chapitre 17.5) ;
- **respecter les limites** du serveur de messagerie pour ne pas être assimilé à du spam.

---

## 7.10.9. Fichiers temporaires et journalisation

- **Fichiers temporaires** : générer les PDF dans `Environ("TEMP")`, puis les supprimer après envoi (`Kill`), ou les conserver dans un dossier d'archive. La manipulation de fichiers via `FileSystemObject` est traitée au chapitre 22.10.
- **Journalisation** : pour la traçabilité, enregistrer dans une table de log qui a reçu quoi et quand (date, destinataire, état, succès/échec). C'est indispensable pour un processus de diffusion en production — voir les chapitres 13.6 (journalisation des erreurs) et 19.5 (journalisation d'événements applicatifs).

---

## 7.10.10. Pièges et bonnes pratiques

- **Pour un état filtré, ouvrir d'abord (caché)** : ni `SendObject` ni `OutputTo` n'acceptent de `WhereCondition` (chapitres 7.4 et 7.8).
- **Préférer l'automation Outlook** pour les envois automatisés, afin d'éviter l'avertissement de sécurité de `SendObject` (chapitre 22.5).
- **Ignorer les destinataires sans données** par un test proactif `DCount` (chapitre 7.9), pour ne pas envoyer de pièces jointes vides.
- **Isoler la gestion d'erreurs par itération** : un échec ne doit jamais interrompre tout le lot (chapitre 13).
- **Refermer systématiquement l'état caché**, y compris en cas d'erreur, via un test `IsLoaded`.
- **Nettoyer les fichiers temporaires** après envoi et gérer les fichiers verrouillés.
- **Tracer les envois** dans une table de log pour la production (chapitres 13.6 et 19.5).
- **Respecter les volumes et valider les adresses** pour éviter rejets et classification en spam.

---

## 7.10.11. Récapitulatif

- Un **état paramétré** produit des sorties multiples à partir d'un état unique, alimenté par un **filtre** (`WhereCondition`, chapitre 7.4) et un **contexte** (`OpenArgs`, chapitre 6.6).
- Deux voies d'envoi : **`DoCmd.SendObject`** (intégré, simple, mais sujet à un avertissement de sécurité) et l'**automation Outlook** (contrôle total, pièce jointe pré-exportée — chapitre 22.5).
- Comme `OutputTo`, `SendObject` n'accepte pas de filtre : pour un envoi **paramétré**, ouvrir l'état filtré (caché) au préalable (chapitre 7.8).
- Le **pattern complet** enchaîne paramétrage → export PDF → e-mail, encapsulé dans une procédure unitaire à gestion d'erreurs, puis répété en **boucle** sur un `Recordset` de destinataires pour l'envoi en lot.
- Pour un processus de production fiable : test proactif des données (chapitre 7.9), gestion d'erreurs **par itération** (chapitre 13), nettoyage des fichiers temporaires (chapitre 22.10) et **journalisation** des envois (chapitres 13.6 et 19.5).

---

> ✅ **Fin du chapitre 7.** L'ensemble du cycle des états est couvert : comprendre leur structure (7.1–7.3), agir sur leurs données et leur rendu (7.4–7.5), composer des états riches (7.6–7.7), puis produire et diffuser les résultats (7.8–7.10).

⏭️ [8. Contrôles et événements](/08-controles-evenements/README.md)
