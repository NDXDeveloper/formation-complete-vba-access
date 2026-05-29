🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 10.3. L'objet Connection — ouverture et fermeture

La section précédente nous a appris à décrire *comment* se connecter, sous la forme d'une chaîne de texte. Il s'agit maintenant de donner vie à cette description : l'objet **`Connection`** est celui qui établit concrètement la session avec la source de données, la maintient ouverte le temps du travail, puis la referme proprement. C'est la pierre angulaire d'ADO — sans connexion ouverte, ni `Recordset` ni `Command` ne peuvent fonctionner.

Maîtriser le cycle ouverture/fermeture est essentiel pour deux raisons : la **fiabilité** (une connexion mal refermée ou un objet mal libéré peut laisser des verrous résiduels et provoquer des erreurs en environnement partagé) et la **performance** (l'ouverture d'une connexion a un coût qu'il faut savoir maîtriser).

---

## Rôle de l'objet Connection

L'objet `Connection` représente une **session ouverte** avec une source de données. Au-delà de simplement « se brancher » sur la base, il joue plusieurs rôles :

Il sert de **support aux autres objets** : un `Recordset` ou un `Command` s'appuient sur une connexion pour dialoguer avec les données. Il **héberge la collection `Errors`**, qui recense les erreurs remontées par le fournisseur (détaillée en section 10.8). Il **pilote les transactions**, via les méthodes `BeginTrans`, `CommitTrans` et `RollbackTrans` (chapitre 14.3). Et il permet, par sa méthode `Execute`, de **lancer directement une commande SQL** sans passer par un objet `Command` distinct.

---

## Créer et ouvrir une connexion

L'instanciation suit le schéma habituel d'ADO, en liaison précoce ou tardive :

```vba
' Liaison précoce (référence « Microsoft ActiveX Data Objects » activée)
Dim cn As ADODB.Connection
Set cn = New ADODB.Connection

' Liaison tardive (aucune référence à activer)
Dim cn As Object
Set cn = CreateObject("ADODB.Connection")
```

L'ouverture proprement dite se fait par la méthode **`Open`**, qui admet deux écritures équivalentes.

**Écriture 1 — la chaîne est passée directement à `Open` :**

```vba
Set cn = New ADODB.Connection
cn.Open "Provider=Microsoft.ACE.OLEDB.12.0;" & _
        "Data Source=" & CurrentProject.Path & "\Donnees.accdb;"
```

**Écriture 2 — la chaîne est affectée à la propriété `ConnectionString`, puis `Open` est appelée sans argument :**

```vba
Set cn = New ADODB.Connection
cn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
                      "Data Source=" & CurrentProject.Path & "\Donnees.accdb;"
cn.CursorLocation = adUseClient   ' configuration AVANT ouverture
cn.Open
```

La seconde écriture n'est pas qu'une question de style : elle est **indispensable dès qu'il faut configurer la connexion avant de l'ouvrir**, comme nous allons le voir.

### L'ordre compte : configurer avant d'ouvrir

C'est l'un des pièges les plus fréquents avec ADO. Plusieurs propriétés de la connexion — au premier rang desquelles **`CursorLocation`** — n'ont d'effet que si elles sont définies **avant** l'appel à `Open`. Une fois la connexion ouverte, elles deviennent en lecture seule ou sont ignorées.

Le tableau suivant récapitule les principales propriétés et le moment où les positionner.

| Propriété | Rôle | Valeur par défaut | À définir avant `Open` ? |
|---|---|---|---|
| `ConnectionString` | Description de la connexion | `""` | Avant, ou passée à `Open` |
| `CursorLocation` | Emplacement du curseur (serveur/client) | `adUseServer` | **Oui, impératif** |
| `ConnectionTimeout` | Délai d'attente d'établissement (s) | `15` | **Oui, impératif** |
| `Mode` | Mode d'accès / de partage | `adModeUnknown` | **Oui, impératif** |
| `Provider` | Fournisseur OLE DB | (selon la chaîne) | **Oui, impératif** |
| `CommandTimeout` | Délai d'attente d'une commande (s) | `30` | Avant l'exécution d'une commande |

> ⚠️ **Erreur classique** : positionner `CursorLocation = adUseClient` *après* `cn.Open` n'a aucun effet. Les recordsets resteront en curseur serveur, et les fonctionnalités déconnectées (section 10.7) ne seront pas disponibles. Configurez toujours **avant** d'ouvrir.

---

## Vérifier l'état : la propriété `State`

La propriété **`State`** renseigne sur l'état courant de la connexion. Sa valeur la plus utile est `adStateOpen`, qui permet de tester si la connexion est effectivement ouverte avant d'agir.

| Constante | Valeur | Signification |
|---|---|---|
| `adStateClosed` | 0 | Connexion fermée |
| `adStateOpen` | 1 | Connexion ouverte |
| `adStateConnecting` | 2 | Connexion en cours d'établissement |
| `adStateExecuting` | 4 | Commande en cours d'exécution |
| `adStateFetching` | 8 | Récupération de lignes en cours |

Ce test est particulièrement important **avant de fermer** : appeler `Close` sur une connexion déjà fermée déclenche une erreur d'exécution. On se prémunit ainsi :

```vba
If cn.State = adStateOpen Then cn.Close
```

---

## Fermer et libérer

Refermer proprement une connexion demande **deux gestes complémentaires**, qu'il ne faut pas confondre :

```vba
cn.Close          ' 1. ferme la session avec la source de données
Set cn = Nothing  ' 2. libère la référence à l'objet en mémoire
```

La méthode **`Close`** met fin à la session : elle relâche la connexion au fichier ou au serveur, libère les verrous éventuels et invalide les recordsets dépendants restés en curseur serveur. L'objet `Connection`, lui, **existe toujours** après `Close` — il peut d'ailleurs être rouvert. L'affectation **`Set cn = Nothing`** libère quant à elle la référence à l'objet, permettant au système de récupérer la mémoire.

Omettre l'un de ces gestes est une source classique de problèmes : une connexion non fermée maintient inutilement des verrous (gênant en multi-utilisateur, chapitre 15), tandis qu'un objet non libéré contribue aux fuites de mémoire dans les applications de longue durée.

> 📌 **Conséquence sur les recordsets** : fermer une connexion ferme aussi les recordsets en curseur serveur qui en dépendent. Les recordsets *déconnectés* (curseur client, `ActiveConnection = Nothing`) survivent en revanche à la fermeture de leur connexion d'origine — c'est tout l'intérêt du mécanisme étudié en section 10.7.

---

## Rappel : le cas particulier de `CurrentProject.Connection`

Tout ce qui précède concerne les connexions que **vous instanciez vous-même** avec `New ADODB.Connection`. La connexion `CurrentProject.Connection`, vue en section 10.2, obéit à une règle différente : elle est **gérée par Access**, déjà ouverte, et **ne doit pas être fermée** par votre code. Vous l'utilisez telle quelle, sans `Open` ni `Close`.

La distinction est simple à retenir : **on referme et on libère ce que l'on a soi-même ouvert et instancié** ; on ne touche pas au cycle de vie d'une connexion qui appartient à Access.

---

## Quelques propriétés utiles à l'ouverture

### Les délais d'attente

La propriété **`ConnectionTimeout`** fixe le nombre de secondes pendant lesquelles ADO tente d'établir la connexion avant d'abandonner (15 secondes par défaut). Elle prend tout son sens face à un serveur distant lent ou momentanément indisponible : un délai trop long fige l'application, un délai trop court échoue prématurément. Pour une base ACE locale, ce réglage est généralement sans enjeu.

La propriété **`CommandTimeout`** régit, elle, le temps d'attente d'une commande déjà lancée (30 secondes par défaut). La valeur `0` signifie « attendre indéfiniment » — à manier avec prudence. Elle s'avère utile pour des requêtes volumineuses dont on sait qu'elles peuvent légitimement dépasser la durée par défaut.

### L'emplacement du curseur

La propriété **`CursorLocation`** détermine si le curseur est géré côté serveur (`adUseServer`, valeur par défaut) ou côté client (`adUseClient`). Ce choix, fondamental, conditionne les types de curseurs disponibles et l'accès aux recordsets déconnectés. Il fait l'objet d'un traitement approfondi en section 10.4, mais retenez dès maintenant qu'il doit être positionné **avant** l'ouverture.

### Le mode d'accès

La propriété **`Mode`** contrôle le partage du fichier (lecture seule, exclusif, partagé…), dans le prolongement du paramètre `Mode` de la chaîne de connexion évoqué en 10.2. Là encore, ce réglage doit précéder l'ouverture, et ses implications en environnement concurrent relèvent du chapitre 15.

---

## Exécuter directement sur la connexion : la méthode `Execute`

L'objet `Connection` ne se contente pas d'établir le lien : il peut **exécuter directement une commande SQL** via sa méthode `Execute`, sans qu'il soit nécessaire de créer un objet `Command` ni d'ouvrir explicitement un `Recordset`. C'est pratique pour les requêtes action :

```vba
Dim lngLignes As Long
cn.Execute "UPDATE Clients SET Actif = True WHERE Region = 'Nord';", lngLignes, adCmdText
Debug.Print lngLignes & " enregistrement(s) modifié(s)."
```

Le deuxième paramètre, ici `lngLignes`, reçoit en sortie le **nombre de lignes affectées**. Le troisième précise la nature de la commande (`adCmdText` pour du SQL brut). Pour une requête `SELECT`, `Execute` renvoie un `Recordset`. Le détail des opérations de lecture et d'écriture est traité en section 10.5, et l'exécution de SQL plus largement au chapitre 11.

---

## Un patron robuste : ouverture, travail, libération garantie

En production, l'ouverture et la fermeture d'une connexion s'inscrivent dans une **structure de gestion d'erreurs** qui garantit la libération des ressources *quoi qu'il arrive* — y compris en cas d'erreur en cours de traitement. Le patron ci-dessous, à adapter à vos besoins, illustre cette discipline (la gestion d'erreurs est approfondie au chapitre 13).

```vba
Public Sub TraiterDonnees()
    Dim cn As ADODB.Connection

    On Error GoTo Gestion

    Set cn = New ADODB.Connection
    cn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
                          "Data Source=" & CurrentProject.Path & "\Donnees.accdb;"
    cn.ConnectionTimeout = 20
    cn.CursorLocation = adUseClient   ' configuration avant ouverture
    cn.Open

    ' ----- Traitement principal -----
    ' (lecture, écriture, exécution de commandes...)
    ' ---------------------------------

Nettoyage:
    On Error Resume Next
    If Not cn Is Nothing Then
        If cn.State = adStateOpen Then cn.Close
        Set cn = Nothing
    End If
    Exit Sub

Gestion:
    MsgBox "Erreur " & Err.Number & " : " & Err.Description, vbCritical, "Connexion"
    Resume Nettoyage
End Sub
```

Trois principes structurent ce patron. La connexion est ouverte **après** l'installation du gestionnaire d'erreurs (`On Error GoTo`), de sorte qu'un échec d'ouverture soit lui-même intercepté. La libération est centralisée dans une **étiquette `Nettoyage`** par laquelle passent obligatoirement le chemin normal *et* le chemin d'erreur. Enfin, on **teste l'état et la nullité** de l'objet avant de le fermer, pour éviter de déclencher une nouvelle erreur dans la phase de nettoyage elle-même.

---

## Réutiliser plutôt que rouvrir

L'ouverture d'une connexion n'est pas gratuite : elle mobilise le fournisseur, vérifie les accès, établit la session. Multiplier les cycles ouverture/fermeture pour des opérations rapprochées pénalise donc inutilement les performances.

La bonne pratique consiste à **ouvrir une connexion une seule fois**, à effectuer l'ensemble des opérations qui en dépendent, puis à la refermer — plutôt que de la rouvrir à chaque requête. Dans les applications qui sollicitent intensément la base, on conserve même parfois une connexion ouverte pour toute la durée d'une session de travail. Cette stratégie de *persistance de connexion*, particulièrement pertinente sur les bases partagées en réseau, est développée au chapitre 18.8.

---

## Connexion, transactions et événements (renvois)

L'objet `Connection` est aussi le point d'ancrage de fonctionnalités traitées ailleurs dans cette formation, et qu'il suffit ici de situer.

Les **transactions** (`BeginTrans` / `CommitTrans` / `RollbackTrans`) permettent de regrouper plusieurs opérations en une unité atomique ; elles font l'objet de la section 14.3. La collection **`Errors`**, qui recense les erreurs du fournisseur de manière plus fine que l'objet `Err` de VBA, est détaillée en section 10.8. Enfin, l'objet `Connection` expose des **événements** (`ConnectComplete`, `Disconnect`, `ExecuteComplete`, `InfoMessage`…) que l'on peut capter en déclarant la connexion `WithEvents` dans un module de classe — un mécanisme abordé au chapitre 16.4.

---

## En résumé

L'objet `Connection` établit et maintient la session avec la source de données. On l'instancie (`New`), on le configure **avant** de l'ouvrir — `CursorLocation`, délais et mode ne tolèrent pas l'inverse —, puis on l'ouvre avec `Open`. On vérifie son état via `State`, on le referme avec `Close` et on libère la référence par `Set … = Nothing` ; ces deux gestes sont complémentaires et tous deux nécessaires. Pour la base courante, on s'appuie sur `CurrentProject.Connection`, que l'on ne ferme jamais. Enfin, un patron de gestion d'erreurs garantit la libération des ressources en toute circonstance, et la réutilisation d'une connexion ouverte préserve les performances.

La connexion étant désormais établie, l'objet qui va réellement transporter les données entre en scène. C'est lui qui occupera l'essentiel des sections suivantes : le `Recordset`, avec ses notions propres de curseur et de verrouillage.


⏭️ [10.4. Objet Recordset ADO — curseurs et modes de verrouillage](/10-ado-access/04-recordset-ado-curseurs.md)
