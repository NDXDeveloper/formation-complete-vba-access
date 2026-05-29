🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 22.7. Appel de Web Services REST depuis VBA Access (MSXML2, WinHttp)

Avec cette section, le chapitre quitte l'automation des applications Office pour ouvrir Access sur le **web**. Appeler un service REST depuis VBA permet de dialoguer avec des services extérieurs — géocodage d'adresses, taux de change, validation de données, services d'emailing — ou avec des microservices internes à l'organisation. C'est le visage le plus contemporain et le plus connecté d'Access : la base de données cesse de travailler en vase clos pour s'insérer dans un écosystème d'API.

Le titre nomme les deux familles de composants employées pour émettre des requêtes HTTP : **MSXML2** et **WinHttp**. La réponse d'un service REST étant le plus souvent au format JSON, son analyse fait l'objet de la section suivante (22.8), à laquelle cette section renvoie.

## Qu'est-ce qu'un appel REST, et pourquoi depuis Access ?

Un service REST s'interroge par de simples **requêtes HTTP** adressées à une URL : on indique une **méthode** (lire, créer, modifier, supprimer), d'éventuels **en-têtes** (format attendu, authentification) et, le cas échéant, un **corps** de requête ; le service renvoie un **code de statut** et un **corps de réponse**, généralement en JSON.

Les usages depuis une application Access sont nombreux : enrichir ou valider une donnée à la saisie, récupérer une information externe pour la stocker, déclencher un traitement sur un service distant, ou encore — comme évoqué à la section 22.5 — **envoyer un courriel via une API web**, alternative moderne au SMTP lorsque ce dernier ne convient pas.

## Les composants : MSXML2 et WinHttp

Trois objets COM, tous disponibles en standard sous Windows et utilisables en liaison tardive (section 2.6), permettent d'émettre des requêtes :

- **`MSXML2.XMLHTTP.6.0`** : le client HTTP « côté client », proche de celui d'un navigateur. Il s'appuie sur les réglages système (proxy, cache) et convient aux usages **interactifs**.
- **`MSXML2.ServerXMLHTTP.6.0`** : la variante « côté serveur », indépendante de ces réglages, plus robuste, dotée de **délais d'attente** configurables. Adaptée aux usages **non surveillés**.
- **`WinHttp.WinHttpRequest.5.1`** : le composant WinHTTP, lui aussi conçu pour les scénarios serveur, robuste, avec gestion fine des délais, de l'authentification et du chiffrement.

En pratique, **`ServerXMLHTTP` et `WinHttp` sont à privilégier** pour leur robustesse, leur maîtrise des délais d'attente, leur indépendance vis-à-vis des réglages d'Internet Explorer et leur bonne prise en charge des protocoles de chiffrement modernes ; `XMLHTTP` se réserve aux cas simples et interactifs. Il convient d'employer les versions récentes (6.0 pour MSXML, 5.1 pour WinHTTP).

## Le schéma commun d'une requête

Les trois composants partagent une **interface quasi identique**, ce qui rend le code aisément transposable de l'un à l'autre. Une requête se déroule toujours en quatre temps :

1. **`Open(méthode, url, async)`** prépare la requête. Le troisième paramètre, `False`, indique un appel **synchrone** — le mode courant en VBA.
2. **`setRequestHeader`** ajoute les en-têtes ; il doit être appelé **après `Open` et avant `Send`**.
3. **`Send([corps])`** émet la requête, éventuellement accompagnée d'un corps.
4. La lecture du résultat via **`Status`** (le code HTTP) et **`responseText`** (le corps de la réponse).

Un point mérite d'être souligné : un appel synchrone **bloque le thread** jusqu'à la réponse, ce qui fige Access exactement comme le ferait `Sleep` (section 22.2). D'où l'importance des délais d'attente, abordés plus bas.

## Une requête GET

La requête la plus simple lit une ressource. On vérifie systématiquement le statut avant d'exploiter la réponse :

```vba
Public Function AppelGET(ByVal url As String) As String
    Dim http As Object

    On Error GoTo Gestion
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")

    http.Open "GET", url, False              ' False = synchrone
    http.setRequestHeader "Accept", "application/json"
    http.Send

    If http.Status = 200 Then
        AppelGET = http.responseText
    Else
        AppelGET = ""
        Debug.Print "Erreur HTTP " & http.Status & " : " & http.statusText
    End If

Sortie:
    Set http = Nothing
    Exit Function

Gestion:
    AppelGET = ""
    Debug.Print "Échec de la requête : " & Err.Description
    Resume Sortie
End Function
```

## Une requête POST avec corps JSON

Pour créer ou soumettre une ressource, on emploie `POST` (ou `PUT`, `PATCH`, `DELETE` selon l'opération), en précisant le type de contenu et en transmettant le corps à `Send`. L'exemple suivant, fondé sur WinHTTP, ajoute des **délais d'attente** et un **jeton d'authentification** :

```vba
Public Function AppelPOST(ByVal url As String, ByVal corpsJSON As String, _
                          Optional ByVal jeton As String = "") As String
    Dim http As Object

    On Error GoTo Gestion
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")

    ' Délais d'attente en ms : résolution, connexion, envoi, réception
    http.SetTimeouts 5000, 5000, 10000, 30000

    http.Open "POST", url, False
    http.setRequestHeader "Content-Type", "application/json"
    http.setRequestHeader "Accept", "application/json"
    If Len(jeton) > 0 Then
        http.setRequestHeader "Authorization", "Bearer " & jeton
    End If
    http.Send corpsJSON

    Select Case http.Status
        Case 200, 201           ' OK ou Créé
            AppelPOST = http.responseText
        Case Else
            AppelPOST = ""
            Debug.Print "Erreur HTTP " & http.Status & " : " & http.statusText
    End Select

Sortie:
    Set http = Nothing
    Exit Function

Gestion:
    AppelPOST = ""
    Debug.Print "Échec de la requête : " & Err.Description
    Resume Sortie
End Function
```

La construction du corps JSON transmis à `Send` relève des techniques de la section 22.8.

## En-têtes et authentification

Les **en-têtes** transportent les métadonnées de la requête : `Accept` (format de réponse souhaité), `Content-Type` (format du corps envoyé), et surtout `Authorization` pour l'authentification.

Les API modernes s'authentifient le plus souvent par **jeton** placé dans l'en-tête `Authorization`, sous la forme `Bearer <jeton>` comme ci-dessus, ou par **clé d'API** dans un en-tête dédié (par exemple `X-API-Key`). L'**authentification basique**, qui transmet un identifiant et un mot de passe encodés en base64, peut quant à elle être gérée directement par WinHTTP via sa méthode `SetCredentials`, ce qui évite d'encoder soi-même les identifiants.

Une précaution de sécurité s'impose ici : les clés et jetons d'API ne doivent pas être codés en clair dans une application distribuée. La compilation en ACCDE protège le code (sections 20.3 et 21.1) sans constituer une garantie absolue ; la gestion des secrets mérite la même prudence que tout élément sensible (section 20.5).

## Délais d'attente et robustesse

Pour un appel non surveillé, **configurer des délais d'attente est essentiel**. Sans eux, un serveur qui ne répond pas fige Access indéfiniment. WinHTTP et ServerXMLHTTP exposent à cet effet une méthode `SetTimeouts` (résolution du nom, connexion, envoi, réception, en millisecondes), illustrée ci-dessus.

S'agissant du **chiffrement**, les API actuelles imposent une connexion HTTPS reposant sur des protocoles de chiffrement récents. WinHTTP et ServerXMLHTTP, sur un Windows à jour, les prennent correctement en charge — ce qui en fait, là encore, l'alternative moderne au composant CDO évoqué à la section 22.5, dont la prise en charge du chiffrement est datée.

## Interpréter le résultat et gérer les erreurs

La gestion des erreurs obéit à une distinction fondamentale, déjà rencontrée à propos des API Windows (section 22.1) :

- Les **échecs de transport** (serveur injoignable, nom non résolu, délai dépassé) déclenchent une **erreur d'exécution VBA**, qu'il faut intercepter par `On Error` (chapitre 13).
- Les **erreurs de niveau HTTP** (4xx, 5xx) ne déclenchent **aucune** erreur VBA : le composant renvoie simplement un **code de statut** qu'il faut examiner. Un appel qui « réussit » techniquement peut donc retourner un statut 404 ou 401 : ne pas vérifier `Status`, c'est risquer de traiter un message d'erreur comme s'il s'agissait de données.

Les codes les plus courants sont 200 (OK), 201 (Créé) et 204 (Pas de contenu) pour les succès ; 400 (requête invalide), 401 (non authentifié), 403 (interdit), 404 (introuvable), 429 (trop de requêtes) et 500 (erreur serveur) pour les échecs. Les exemples ci-dessus combinent les deux niveaux de protection : `On Error` pour le transport, examen de `Status` pour le HTTP.

Une dernière subtilité concerne la **construction de l'URL** : les valeurs de paramètres comportant des espaces ou des caractères spéciaux doivent être **encodées pour l'URL**. VBA ne disposant pas, dans Access, d'une fonction d'encodage intégrée, on écrit à cette fin une petite fonction utilitaire, sous peine de requêtes mal formées.

## Articulation avec le reste du chapitre et de la formation

Cette section ouvre le volet connectivité du chapitre et s'articule avec plusieurs autres :

- L'**analyse du JSON** renvoyé par les services est traitée à la section suivante (22.8), prolongement naturel de celle-ci.
- L'appel REST constitue l'**alternative moderne** au SMTP pour l'envoi de courriels (section 22.5).
- La **liaison tardive** des composants relève de la section 2.6.
- La **gestion d'erreurs**, avec sa distinction entre échec de transport et erreur HTTP, relève du chapitre 13, et la **protection des secrets** (clés, jetons) des sections 20.3, 20.5 et 21.1.

## Points de vigilance

Plusieurs erreurs reviennent fréquemment dans l'appel de services REST :

- **Ne pas vérifier le code de statut**, et traiter une réponse 4xx ou 5xx comme des données valides.
- **Omettre les délais d'attente** sur un appel non surveillé, exposant Access à un blocage indéfini.
- **Appeler un service synchrone sur une opération lente**, figeant l'interface le temps de la réponse.
- **Placer `setRequestHeader` avant `Open`**, alors qu'il doit l'être après `Open` et avant `Send`.
- **Coder en clair une clé ou un jeton d'API** dans une application distribuée.
- **Oublier d'encoder les paramètres d'URL**, produisant des requêtes mal formées.
- **Recourir à `XMLHTTP` pour un usage serveur** alors que `ServerXMLHTTP` ou `WinHttp` sont plus robustes.

⏭️ [22.8. Parsing JSON en VBA (sans bibliothèque externe)](/22-api-windows-integration-office/08-parsing-json.md)
