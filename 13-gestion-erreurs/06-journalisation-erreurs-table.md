🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 13.6. Journalisation des erreurs dans une table Access

## Introduction

Intercepter une erreur et l'afficher à l'écran suffit à protéger l'utilisateur sur le moment, mais ne laisse aucune trace exploitable une fois la boîte de dialogue refermée. Or, en production, l'instant où l'erreur survient est précisément celui où le développeur est absent. L'utilisateur, lui, retient rarement le numéro et la description exacts, et son compte rendu se résume souvent à « ça a planté ». Pour diagnostiquer sérieusement une application déployée, il faut donc **conserver une trace** de chaque incident, indépendamment de l'utilisateur et du moment.

La journalisation répond à ce besoin : elle enregistre, à chaque erreur, l'ensemble des informations utiles dans un emplacement durable — ici, une table Access. Cette section décrit la conception d'une telle table, la manière d'y écrire de façon fiable, et les précautions propres à l'environnement Access, notamment l'épineuse question de l'emplacement du journal dans une architecture partagée. La procédure de journalisation construite ici constituera par ailleurs la brique centrale du module de gestion centralisée présenté à la section [13.7](/13-gestion-erreurs/07-module-centralise-erreurs.md).

## Pourquoi journaliser les erreurs

La journalisation apporte une valeur que ni le message à l'écran ni le débogage en atelier ne peuvent offrir.

Elle permet d'abord de diagnostiquer les incidents survenus **sur le poste de l'utilisateur**, à distance, sans avoir à reproduire le contexte exact qui a déclenché l'erreur. Elle rend ensuite visibles les erreurs **intermittentes** — celles qui ne se manifestent que dans certaines conditions, parfois liées à un poste, un horaire ou un jeu de données particulier — que l'on ne parvient jamais à reproduire en environnement de développement. Elle construit enfin un **historique** : en agrégeant les entrées, on distingue l'incident isolé du problème récurrent, on mesure la fréquence d'une défaillance, on repère les procédures les plus fragiles et l'on hiérarchise les corrections en fonction de leur impact réel.

En somme, le journal transforme des incidents ponctuels et fuyants en données objectives, datées et analysables — la matière première de toute démarche de maintenance sérieuse.

## Concevoir la table de journal

La qualité du diagnostic dépend directement des informations consignées. Une table de journal bien conçue capture, pour chaque erreur, le **quoi**, le **où**, le **qui** et le **quand**.

| Champ | Type | Rôle |
|---|---|---|
| `ID_Erreur` | NuméroAuto | Clé primaire |
| `DateHeure` | Date/Heure | Horodatage de l'incident (`Now`) |
| `Utilisateur` | Texte court | Identifiant de l'utilisateur (`Environ("Username")`) |
| `Poste` | Texte court | Nom de la machine (`Environ("Computername")`) |
| `NumeroErreur` | Numérique (Entier long) | `Err.Number` |
| `Description` | Texte long | `Err.Description` (peut être volumineuse) |
| `Source` | Texte court | `Err.Source` |
| `Procedure` | Texte court | Nom de la procédure où l'erreur est survenue |
| `Module` | Texte court | Nom du module (facultatif) |
| `Ligne` | Numérique | Numéro de ligne via `Erl` (si numérotation activée) |
| `VersionApp` | Texte court | Version de l'application (facultatif) |
| `Contexte` | Texte long | Informations complémentaires libres (facultatif) |

Quelques remarques sur ces choix. Le champ `Description` est de type **Texte long** (l'ancien type « Mémo »), car certaines descriptions — en particulier les erreurs ODBC complètes vues à la section [13.5](/13-gestion-erreurs/05-erreurs-dao-ado.md) — dépassent largement la limite d'un champ texte court. Les champs `Utilisateur` et `Poste` situent l'incident dans l'environnement réel, ce qui est précieux pour repérer un problème localisé à un poste précis. Le champ `Procedure` localise l'erreur dans le code ; combiné à `Ligne` (alimenté par `Erl` lorsque la numérotation des lignes est en place, voir section [13.2](/13-gestion-erreurs/02-on-error-goto.md)), il permet de remonter à l'instruction exacte. Le champ `Contexte`, enfin, accueille toute information utile au cas par cas : la valeur d'un identifiant en cours de traitement, l'opération tentée, etc.

## Le principe directeur : la journalisation ne doit jamais échouer

Avant d'écrire la moindre ligne de code, il faut poser un principe absolu : **la journalisation ne doit jamais, sous aucun prétexte, interrompre l'application ni masquer l'erreur d'origine.**

Ce principe a une conséquence concrète. La procédure de journalisation manipule elle-même la base de données — elle ouvre un recordset, écrit un enregistrement — et ces opérations peuvent à leur tour échouer : table absente, base en lecture seule, connexion au back-end perdue. Si la journalisation déclenchait une erreur non gérée, elle aggraverait la situation au lieu de l'éclairer, et risquerait de remplacer le message de l'erreur réelle par le sien. Une routine de journalisation doit donc **gérer ses propres erreurs de façon autonome**, et en cas d'échec, se replier silencieusement plutôt que propager quoi que ce soit. Mieux vaut une erreur non journalisée qu'une application déstabilisée par sa propre journalisation.

## Écrire dans la table de façon fiable

### L'approche Recordset, recommandée

Deux voies permettent d'insérer un enregistrement dans la table de journal : une requête `INSERT` en SQL, ou un recordset DAO via `AddNew`/`Update`. La seconde est nettement préférable pour la journalisation.

La raison tient à la nature des données enregistrées. Les descriptions d'erreur contiennent fréquemment des apostrophes (« l'enregistrement », « n'existe pas »), des guillemets, voire des retours à la ligne. Une requête `INSERT` construite par concaténation devrait échapper soigneusement ces caractères, faute de quoi elle provoquerait elle-même une erreur de syntaxe — au pire moment, puisqu'on cherche justement à journaliser une erreur. L'approche Recordset **affecte directement les valeurs aux champs**, sans aucune concaténation ni échappement : le texte est transmis tel quel, quels que soient les caractères qu'il contient. Elle est donc à la fois plus simple et plus robuste.

### Pourquoi éviter l'INSERT SQL naïf

Un `INSERT` construit ainsi serait fragile :

```vba
' À ÉVITER pour la journalisation : la description casse la requête
CurrentDb.Execute "INSERT INTO T_JournalErreurs (Description) " & _
                  "VALUES ('" & description & "')"
```

Si `description` vaut `L'objet n'existe pas`, l'apostrophe ferme prématurément la chaîne SQL et déclenche une erreur de syntaxe. Le doublement systématique des apostrophes serait nécessaire (technique détaillée à la section 11.5 sur le SQL paramétré), mais l'approche Recordset rend cette précaution inutile et constitue, ici, le bon choix par défaut.

## Une procédure de journalisation réutilisable

Voici une procédure complète appliquant ces principes : écriture par recordset, gestion autonome des erreurs, repli en cas d'échec, et réception des informations par paramètres.

```vba
Public Sub JournaliserErreur(ByVal numero As Long, _
                             ByVal description As String, _
                             ByVal source As String, _
                             ByVal procedure As String, _
                             Optional ByVal ligne As Long = 0, _
                             Optional ByVal contexte As String = "")

    ' La journalisation gère ses propres erreurs et n'en propage aucune.
    On Error GoTo EchecJournalisation

    Dim rst As DAO.Recordset
    Set rst = CurrentDb.OpenRecordset("T_JournalErreurs", _
                                      dbOpenDynaset, dbAppendOnly)

    rst.AddNew
    rst!DateHeure = Now()
    rst!Utilisateur = Environ$("Username")
    rst!Poste = Environ$("Computername")
    rst!NumeroErreur = numero
    rst!Description = description        ' aucun échappement requis
    rst!Source = source
    rst!Procedure = procedure
    rst!Ligne = ligne
    rst!Contexte = contexte
    rst.Update

    rst.Close
    Set rst = Nothing
    Exit Sub

EchecJournalisation:
    ' Repli : on tente le fichier texte ; en dernier recours, on abandonne.
    On Error Resume Next
    JournaliserDansFichier numero, description, source, procedure
    If Not rst Is Nothing Then rst.Close
    Set rst = Nothing
End Sub
```

L'option `dbAppendOnly` ouvre le recordset en mode ajout seul, sans charger les enregistrements existants : la journalisation est ainsi plus rapide et plus légère, ce qui compte d'autant plus que la table grossit avec le temps. Le gestionnaire `EchecJournalisation` n'affiche rien et ne lève rien : il tente une dernière sauvegarde dans un fichier, libère les ressources, puis rend la main en silence.

## Le repli vers un fichier texte

Lorsque l'écriture en table est impossible — typiquement parce que la connexion au back-end est précisément ce qui a échoué —, un repli vers un simple fichier texte garantit que l'information n'est pas totalement perdue :

```vba
Private Sub JournaliserDansFichier(ByVal numero As Long, _
                                   ByVal description As String, _
                                   ByVal source As String, _
                                   ByVal procedure As String)
    On Error Resume Next            ' ce repli ne doit jamais échouer non plus

    Dim numFichier As Integer
    Dim chemin As String

    chemin = Environ$("TEMP") & "\Journal_Erreurs.log"
    numFichier = FreeFile

    Open chemin For Append As #numFichier
    Print #numFichier, Format$(Now(), "yyyy-mm-dd hh:nn:ss") & vbTab & _
          Environ$("Username") & vbTab & numero & vbTab & _
          procedure & vbTab & source & vbTab & description
    Close #numFichier
End Sub
```

Le fichier est écrit dans le dossier temporaire de l'utilisateur — un emplacement presque toujours accessible en écriture — au format délimité par tabulations, aisément lisible ou importable ultérieurement. Comme la routine principale, ce repli est intégralement placé sous `On Error Resume Next` : s'il échoue à son tour, l'application n'en sera pas affectée.

## Où placer la table de journal ?

Dans une application Access partagée, organisée en front-end et back-end (chapitre 15), une question décisive se pose : **dans quel fichier la table de journal doit-elle résider ?** Le choix n'est pas neutre, car il oppose deux qualités difficilement conciliables.

Placer la table dans le **back-end** centralise les incidents : l'administrateur dispose d'un journal unique rassemblant les erreurs de tous les utilisateurs, idéal pour l'analyse globale. Mais cette solution souffre d'une faiblesse majeure : si l'erreur à journaliser est justement une **perte de connexion au back-end**, l'écriture du journal échouera elle aussi. Or les erreurs de connexion comptent précisément parmi celles que l'on souhaite le plus tracer.

Placer la table dans le **front-end** — donc dans la copie locale de chaque utilisateur — rend la journalisation **robuste face aux défaillances réseau**, puisqu'elle n'en dépend pas. En contrepartie, les journaux sont dispersés sur autant de postes que d'utilisateurs, ce qui complique la consolidation.

En pratique, on combine souvent ces approches plutôt que de choisir : journalisation dans le back-end lorsqu'il est accessible, et repli automatique vers une table locale ou un fichier texte dans le cas contraire. Le mécanisme de repli vers fichier présenté ci-dessus joue exactement ce rôle de filet de sécurité, et constitue à lui seul une réponse pragmatique au dilemme.

## Passer les informations en paramètres

On remarquera que la procédure `JournaliserErreur` reçoit le numéro, la description et la source **en paramètres**, plutôt que de lire directement l'objet `Err`. Ce choix n'est pas anodin : il découle de la règle établie à la section [13.4](/13-gestion-erreurs/04-objet-err.md). Au moment où la journalisation s'exécute, l'objet `Err` a très probablement déjà été réinitialisé — ne serait-ce que par l'instruction `On Error` interne à la routine de journalisation elle-même. C'est donc au **gestionnaire appelant** de capturer les valeurs de `Err` dès son entrée, puis de les transmettre à `JournaliserErreur`. Cette séparation garantit que l'information consignée correspond bien à l'erreur réelle, et non à un objet `Err` vidé entre-temps.

## Maintenance du journal

Une table de journal grossit indéfiniment si rien ne la régule. Au fil des mois, elle peut atteindre une taille qui pèse sur la base et sur les performances. Il convient donc de prévoir une **purge** ou un **archivage** périodique : suppression des entrées au-delà d'une certaine ancienneté, ou export régulier vers un fichier d'archive avant épuration. Cette opération de maintenance, simple à automatiser, évite que l'outil de diagnostic ne devienne lui-même une source de problèmes (la gestion de la limite de 2 Go d'un fichier Access est abordée au chapitre 18).

## Données sensibles : ce qu'il ne faut pas journaliser

Un journal d'erreurs peut être consulté par plusieurs personnes et conservé longtemps : il ne doit donc jamais contenir d'informations sensibles. Le champ `Contexte`, tentant pour y verser « tout ce qui pourrait servir », appelle une vigilance particulière. On y consigne des **identifiants** et des **références** (le numéro d'un enregistrement, le nom d'une opération), mais jamais de **secrets** ni de **données personnelles** au sens strict — mots de passe, données nominatives complètes, informations confidentielles. Journaliser de quoi diagnostiquer ne signifie pas journaliser de quoi compromettre.

## En résumé

La journalisation conserve une trace durable et analysable de chaque erreur, là où le message à l'écran s'efface dès qu'il est lu. Une table bien conçue capture le numéro, la description et la source de l'erreur, mais aussi sa localisation dans le code (procédure, ligne) et dans l'environnement (utilisateur, poste, horodatage). Le principe directeur est intangible : **la journalisation ne doit jamais échouer ni masquer l'erreur d'origine** ; elle gère ses propres erreurs et se replie en silence le cas échéant.

L'écriture par recordset (`AddNew`/`Update`) est préférable à un `INSERT` SQL, car elle évite tout problème d'échappement avec des descriptions riches en caractères spéciaux. Dans une architecture partagée, l'emplacement de la table oppose la centralisation (back-end) à la robustesse (front-end), un dilemme que résout efficacement un repli automatique vers un fichier texte. Les informations sont transmises en paramètres pour préserver leur intégrité face à la réinitialisation de `Err`, le journal est purgé régulièrement, et l'on s'abstient d'y consigner la moindre donnée sensible.

Cette procédure de journalisation, robuste et autonome, ne demande plus qu'à être orchestrée. C'est l'objet de la section suivante, qui l'intègre dans un module de gestion centralisée des erreurs, appelable de façon uniforme depuis l'ensemble de l'application.

---


⏭️ [13.7. Module de gestion centralisée des erreurs](/13-gestion-erreurs/07-module-centralise-erreurs.md)
