🔝 Retour au [Sommaire](/SOMMAIRE.md)

# C. Codes d'erreur courants Access / DAO / ADO

Cette annexe recense les codes d'erreur les plus fréquemment rencontrés en VBA Access. Elle complète le **chapitre 13** (gestion des erreurs) : ce dernier explique les structures de capture, tandis que cette annexe sert de table de diagnostic pour identifier rapidement une erreur et sa cause probable.

Il s'agit d'une **sélection des erreurs courantes**, pas d'un catalogue exhaustif. La colonne « Cause fréquente » indique l'origine la plus probable, qui n'est pas toujours la seule possible.

## Comment lire un code d'erreur

En VBA, l'erreur est portée par l'objet `Err` :

```vba
MsgBox "Erreur " & Err.Number & " : " & Err.Description & " (source : " & Err.Source & ")"
```

Trois points essentiels :

- **Un même numéro peut avoir des sens différents selon la couche** qui le lève (VBA, Access, moteur ACE/DAO, fournisseur ADO). Le contexte de l'opération en cours est déterminant.
- **DAO et ADO maintiennent leur propre collection `Errors`.** Une seule opération peut générer plusieurs erreurs ; `Err.Number` n'en reflète qu'une. Pour le détail, parcourir `DBEngine.Errors` (DAO) ou `Connection.Errors` (ADO) — chapitres 13.5 et 10.8.
- **La fonction `AccessError(numéro)`** renvoie le texte d'une erreur Access/DAO **sans la déclencher**, utile pour construire des messages ou consulter une description.

```vba
Debug.Print AccessError(3022)   ' Affiche le message associé au code 3022
```

---

## Erreurs VBA d'exécution

Les erreurs classiques du langage, indépendantes d'Access.

| N° | Message (résumé) | Cause fréquente |
|---|---|---|
| 5 | Argument ou appel de procédure incorrect | Valeur d'argument hors plage, fonction mal employée |
| 6 | Dépassement de capacité (Overflow) | Affectation d'un nombre trop grand pour le type (ex. `Integer` > 32767) |
| 7 | Mémoire insuffisante | Traitement trop volumineux, fuite d'objets non libérés |
| 9 | L'indice n'appartient pas à la sélection | Indice de tableau ou de collection hors limites |
| 11 | Division par zéro | Dénominateur nul ou Null converti en 0 |
| 13 | Type incompatible (Type mismatch) | Affectation entre types non convertibles ; champ `Null` lu dans une variable typée |
| 53 | Fichier introuvable | Chemin ou nom de fichier erroné |
| 70 | Autorisation refusée | Fichier en lecture seule, ouvert ailleurs, ou droits insuffisants |
| 76 | Chemin d'accès introuvable | Dossier inexistant (souvent chemin UNC ou lecteur réseau) |
| 91 | Variable objet ou variable de bloc `With` non définie | Objet utilisé sans `Set` préalable, ou `Nothing` |
| 94 | Utilisation incorrecte de Null | Opération sur une valeur `Null` (concaténation, calcul) sans test `IsNull`/`Nz` |
| 380 | Valeur de propriété non valide | Valeur refusée par une propriété (ex. couleur, énumération) |
| 424 | Objet requis | Référence à un objet inexistant ou mal nommé |
| 429 | Le composant ActiveX ne peut pas créer l'objet | `CreateObject` sur une bibliothèque absente ou non enregistrée |
| 438 | Propriété ou méthode non gérée par cet objet | Membre inexistant pour ce type d'objet (souvent faute de frappe ou mauvais type) |
| 457 | Cette clé est déjà associée à un élément de cette collection | Ajout d'une clé en double dans une `Collection` |
| 462 | Le serveur distant n'existe pas ou n'est pas disponible | Référence à une application Automation fermée (Excel/Word/Outlook libéré trop tôt) |

---

## Erreurs Access (application)

Erreurs propres à l'application Access, généralement dans la plage 2xxx.

| N° | Message (résumé) | Cause fréquente |
|---|---|---|
| 2046 | La commande ou l'action n'est pas disponible actuellement | `RunCommand`/`DoCmd` appelé dans un contexte inadapté |
| 2105 | Impossible d'atteindre l'enregistrement spécifié | `GoToRecord acNewRec` interdit, ou enregistrement hors plage |
| 2110 | Impossible de placer le focus sur le contrôle | `SetFocus` sur un contrôle invisible, désactivé ou inexistant |
| 2185 | Référence à une propriété d'un contrôle sans focus | Lecture de `.Text` alors que le contrôle n'a pas le focus (utiliser `.Value`) |
| 2237 | Le texte saisi n'est pas un élément de la liste | Valeur absente d'une liste déroulante à liste imposée |
| 2424 | Nom de champ, contrôle ou propriété introuvable dans l'expression | Faute de frappe dans un nom de contrôle/champ |
| 2427 | L'expression saisie n'a pas de valeur | Référence à un contrôle/sous-formulaire non chargé |
| 2448 | Impossible d'affecter une valeur à cet objet | Affectation à une propriété en lecture seule ou dans un mauvais contexte |
| 2450 | Formulaire introuvable (référencé dans le code) | `Forms!Nom` alors que le formulaire n'est pas ouvert |
| 2465 | Champ introuvable dans l'expression | Nom de champ inexistant dans la source du formulaire/état |
| 2467 | L'expression fait référence à un objet fermé ou inexistant | Objet refermé entre-temps |
| 2474 | L'expression requiert que le contrôle soit dans la fenêtre active | Usage de `Screen.ActiveControl` hors fenêtre active |
| 2489 | L'objet n'est pas ouvert | Manipulation d'un formulaire/état non ouvert |
| 2501 | L'action … a été annulée | **Attendu** quand `Cancel = True` annule un `OpenForm`/`OpenReport` (ex. événement `NoData`) — à intercepter et ignorer |

---

## Erreurs DAO / moteur ACE

Erreurs du moteur de base de données (plage 3xxx), parmi les plus fréquentes en accès aux données.

### Données et Recordset

| N° | Message (résumé) | Cause fréquente |
|---|---|---|
| 3020 | `Update` ou `CancelUpdate` sans `AddNew` ni `Edit` | Oubli d'appeler `.Edit` ou `.AddNew` avant `.Update` |
| 3021 | Aucun enregistrement courant | Lecture d'un champ alors que le Recordset est vide ou positionné en EOF/BOF |
| 3022 | Valeurs en double (index, clé primaire ou relation) | Insertion/modification créant un doublon sur une contrainte d'unicité |
| 3027 | Mise à jour impossible : base ou objet en lecture seule | Recordset `Snapshot`, requête non modifiable, base en lecture seule |
| 3163 | Le champ est trop petit pour les données | Texte dépassant la taille (`FieldSize`) du champ |
| 3167 | L'enregistrement est supprimé | Accès à un enregistrement supprimé entre-temps |

### SQL et requêtes

| N° | Message (résumé) | Cause fréquente |
|---|---|---|
| 3061 | Trop peu de paramètres. N attendu(s) | **Très courant** : nom de champ inconnu dans un SQL dynamique, ou texte non encadré de guillemets — le moteur le prend pour un paramètre |
| 3075 | Erreur de syntaxe dans l'expression de requête | SQL mal formé (apostrophe, parenthèse, opérateur) |
| 3078 | Table ou requête source introuvable | Nom de table/requête erroné dans le SQL |
| 3079 | Le champ peut faire référence à plusieurs tables | Champ ambigu dans une jointure (préfixer par le nom de table) |
| 3134 | Erreur de syntaxe dans l'instruction INSERT INTO | Liste de champs/valeurs incohérente |
| 3144 | Erreur de syntaxe dans l'instruction UPDATE | Clause `SET` malformée |

### Ouverture et fichier

| N° | Message (résumé) | Cause fréquente |
|---|---|---|
| 3024 | Fichier introuvable | Base back-end déplacée ou chemin invalide |
| 3049 | Impossible d'ouvrir la base de données | Fichier corrompu ou format non reconnu |
| 3051 | Le moteur ne peut ouvrir/écrire dans le fichier | Base déjà ouverte en exclusif, ou droits insuffisants sur le dossier |
| 3343 | Format de base de données non reconnu | Version de fichier incompatible (.mdb ancien, fichier endommagé) |

### Verrouillage multi-utilisateur

| N° | Message (résumé) | Cause fréquente |
|---|---|---|
| 3186 | Enregistrement verrouillé par un autre utilisateur | Conflit d'enregistrement sur le réseau (chapitre 15.4) |
| 3187 | Lecture impossible : actuellement verrouillé | Verrou posé par une autre session |
| 3188 | Mise à jour impossible : verrouillé par une autre session de cette machine | Deux instances locales ouvrent la même donnée |
| 3197 | Le moteur a interrompu le traitement (modification simultanée) | Conflit d'écriture entre deux utilisateurs |
| 3260 | Mise à jour impossible : verrouillé par l'utilisateur … | Verrou pessimiste en cours côté autre poste |

### ODBC / serveur lié

| N° | Message (résumé) | Cause fréquente |
|---|---|---|
| 3146 | Échec de l'appel ODBC | Erreur côté serveur lié (SQL Server) ; lire les détails dans `Errors` |
| 3151 | Échec de la connexion ODBC | Chaîne de connexion, droits, ou serveur indisponible (chapitres 11.9 et 23.4) |

---

## Erreurs ADO

ADO renvoie souvent des **HRESULT** (grands nombres négatifs ou hexadécimaux), dépendants du fournisseur OLE DB. Le détail se lit dans la collection `Connection.Errors` (chapitre 10.8).

### Erreurs ADO propres (objets fermés/ouverts)

| N° | Message (résumé) | Cause fréquente |
|---|---|---|
| 3219 | Opération non autorisée dans ce contexte | Méthode incompatible avec l'état/type du curseur |
| 3251 | L'objet ou le fournisseur ne peut effectuer l'opération | Curseur en lecture seule, fonctionnalité non gérée par le fournisseur |
| 3265 | Élément introuvable dans la collection | Nom de champ/paramètre erroné (`rs.Fields("...")`) |
| 3704 | Opération non autorisée : l'objet est fermé | Recordset/Connection utilisé sans `Open` préalable |
| 3705 | Opération non autorisée : l'objet est ouvert | Réouverture d'un objet déjà ouvert |
| 3709 | La connexion est fermée ou non valide | Connexion perdue ou non initialisée |

### HRESULT fréquents du fournisseur ACE/OLE DB

| Hexadécimal | Décimal | Signification |
|---|---|---|
| `0x80040E14` | -2147217900 | Erreur de syntaxe SQL |
| `0x80040E37` | -2147217865 | Table ou objet introuvable |
| `0x80040E4D` | -2147217843 | Échec d'authentification (login) |
| `0x80004005` | -2147467259 | Erreur non spécifiée (générique OLE DB) |

---

## Points de vigilance

> ⚠️ **Le même numéro change de sens selon la couche.** Toujours considérer l'opération en cours et, pour DAO/ADO, consulter la collection `Errors` plutôt que `Err.Number` seul.

> 💡 **Erreur 3061 (« Trop peu de paramètres »).** C'est rarement un vrai paramètre manquant : le plus souvent, un nom de champ est mal orthographié dans un SQL dynamique, ou une valeur texte n'est pas encadrée de guillemets. Le moteur, ne reconnaissant pas le jeton, le prend pour un paramètre attendu (chapitre 11.6).

> 💡 **Erreur 3021 (« Aucun enregistrement courant »).** Tester systématiquement `If Not (rs.EOF And rs.BOF) Then …` avant de lire un champ, après l'ouverture d'un Recordset (chapitre 9.4).

> 💡 **Erreur 3022 (doublon).** Vérifier les contraintes d'unicité (clé primaire, index unique, relation) avant l'insertion, plutôt que de réagir après coup (chapitre 14.6).

> 💡 **Erreur 2501 (action annulée).** C'est un comportement **normal** lorsqu'un événement annule l'ouverture (`Cancel = True` dans `Open`, ou état sans données via `NoData`). L'intercepter explicitement et la traiter comme non bloquante (chapitres 7.9 et 8.12).

> 💡 **Erreurs 91 / 424 / 438.** Le trio le plus fréquent : objet non instancié (`Set` oublié), nom mal orthographié, ou membre inexistant pour ce type d'objet. Première piste : vérifier l'orthographe et le type de l'objet manipulé.

> 💡 **Erreurs de verrouillage (3186, 3187, 3188, 3197, 3260).** Symptomatiques d'un environnement multi-utilisateur. Prévoir une nouvelle tentative temporisée (retry) et un message clair (chapitres 15.3 et 15.4).

---

> 💡 **À retenir.** Identifier d'abord la **couche** qui lève l'erreur (langage VBA, application Access, moteur ACE/DAO, fournisseur ADO) oriente immédiatement le diagnostic. Pour toute erreur absente de cette sélection, `AccessError(numéro)` fournit le libellé d'une erreur Access/DAO, et les collections `Errors` de DAO et ADO donnent le détail complet d'une opération ayant échoué.

⏭️ [D. Chaînes de connexion ADO/ODBC pour les SGBD courants](/annexes/d-chaines-connexion-ado-odbc.md)
