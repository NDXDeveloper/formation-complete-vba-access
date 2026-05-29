🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 18.7. Compactage automatique de la base par code

Une base Access grossit avec le temps, même lorsque la quantité de données n'augmente pas. Le compactage est l'opération de maintenance qui récupère cet espace perdu, réorganise les données et les index, et entretient ainsi à la fois la taille du fichier et la performance. L'automatiser par code permet de l'intégrer à la vie de l'application sans dépendre d'une intervention manuelle — à condition de composer avec une contrainte forte : on ne compacte pas une base ouverte.

## Pourquoi compacter

Le gonflement (bloat) provient de plusieurs sources : les enregistrements supprimés laissent un espace qui n'est pas automatiquement repris, les modifications répétées fragmentent les données, les requêtes créent des objets temporaires, et certaines options (comme la correction automatique des noms) ajoutent de la surcharge. Le fichier croît alors même que les données diminuent.

Le compactage répare cela sur trois plans. Il **récupère l'espace** inutilisé et réduit la taille du fichier — ce qui aide à rester sous la limite des 2 Go (voir la [section 18.9](09-limites-taille-2go.md)). Il **défragmente** les données et reconstruit les index, améliorant les temps d'accès. Et il **rafraîchit les statistiques** de l'optimiseur tout en recompilant les plans des requêtes sauvegardées (voir la [section 18.6](06-indexation-performances-sql.md)). L'opération « Compacter et réparer » corrige en outre les petites corruptions.

## La contrainte fondamentale : on ne compacte pas une base ouverte

Le compactage exige un accès **exclusif** au fichier. On ne peut donc pas compacter directement, de l'intérieur, la base actuellement ouverte : le fichier est en cours d'utilisation. Cette contrainte dicte les approches possibles :

- compacter une **autre** base — typiquement le back-end de données depuis le front-end ;
- s'en remettre au **compactage à la fermeture** pour le fichier courant ;
- ou faire compacter le front-end par un **lanceur externe**.

## Compacter une autre base : `CompactRepair` et `CompactDatabase`

Deux méthodes compactent un fichier fermé vers un nouveau fichier.

`Application.CompactRepair` est la méthode Access ; elle renvoie un booléen indiquant le succès.

```vba
Dim ok As Boolean
ok = Application.CompactRepair( _
        "C:\Data\Donnees.accdb", _
        "C:\Data\Donnees_compact.accdb")
If ok Then Debug.Print "Compactage réussi"
```

`DBEngine.CompactDatabase` est l'équivalent DAO ; la destination ne doit pas exister au préalable. Ses arguments optionnels permettent notamment de fixer la langue, des options (version, chiffrement) ou un mot de passe.

```vba
' La destination ne doit pas exister
DBEngine.CompactDatabase "C:\Data\Donnees.accdb", "C:\Data\Donnees_compact.accdb"
```

Dans les deux cas, la base source ne doit pas être ouverte.

> ℹ️ `DBEngine` et `Workspace` sont présentés au [chapitre 4.5](../04-modele-objet-access/05-dbengine-workspace-database.md) ; le chiffrement par mot de passe au [chapitre 20.2](../20-securite-protection/02-protection-mot-de-passe.md).

## Compacter un fichier « sur place » : le motif d'échange

Comme ces méthodes produisent un **nouveau** fichier, compacter une base à son emplacement d'origine suppose un échange : compacter vers un temporaire, sauvegarder l'original, puis remplacer. Une gestion d'erreurs soigneuse et une copie de sécurité sont indispensables.

```vba
Public Function CompacterFichier(ByVal chemin As String) As Boolean
    Dim fso As Object, tmp As String, sauvegarde As String
    Set fso = CreateObject("Scripting.FileSystemObject")

    tmp = fso.GetParentFolderName(chemin) & "\~compact_" & fso.GetFileName(chemin)
    sauvegarde = chemin & ".bak"

    On Error GoTo Echec

    If fso.FileExists(tmp) Then fso.DeleteFile tmp     ' résidu éventuel

    ' 1. Compacter vers un fichier temporaire
    If Not Application.CompactRepair(chemin, tmp) Then GoTo Echec

    ' 2. Conserver l'original en sauvegarde, puis le remplacer
    If fso.FileExists(sauvegarde) Then fso.DeleteFile sauvegarde
    fso.MoveFile chemin, sauvegarde
    fso.MoveFile tmp, chemin

    CompacterFichier = True
    Exit Function

Echec:
    ' En cas d'échec, l'original (ou sa sauvegarde) reste exploitable
    CompacterFichier = False
End Function
```

Cette fonction vise un fichier **fermé** ; elle ne peut donc pas s'appliquer à la base courante.

> ℹ️ Le `FileSystemObject` est traité au [chapitre 22.10](../22-api-windows-integration-office/10-filesystemobject.md) ; la gestion robuste des erreurs au [chapitre 13](../13-gestion-erreurs/README.md).

## Compacter le back-end depuis le front-end

C'est le scénario d'automatisation le plus courant : le front-end compacte le fichier de données, à condition qu'**aucun autre utilisateur** ne l'ait ouvert. En environnement partagé, cette exclusivité doit être garantie — opération réservée à une plage horaire creuse, ou précédée d'une vérification qu'aucune session n'est active. Le compactage ne doit jamais s'effectuer pendant que d'autres utilisateurs sont connectés.

> ℹ️ La détection des sessions (fichier de verrouillage, table de sessions) est traitée au [chapitre 15.8](../15-multi-utilisateurs/08-gestion-sessions-utilisateurs.md) ; l'architecture front-end / back-end au [chapitre 15](../15-multi-utilisateurs/README.md).

## Déclencher selon un seuil de taille

Plutôt que de compacter systématiquement, on peut conditionner l'opération à la taille du fichier — par exemple au démarrage de l'application.

```vba
If FileLen("C:\Data\Donnees.accdb") > 500& * 1024 * 1024 Then   ' > ~500 Mo
    CompacterFichier "C:\Data\Donnees.accdb"
End If
```

## Le compactage à la fermeture

L'option « Compacter lors de la fermeture » (Compact on Close) compacte automatiquement la base courante à sa fermeture. C'est l'approche la plus simple, activable par code.

```vba
Application.SetOption "Auto Compact", True
```

Sa simplicité a une contrepartie : l'opération peut échouer discrètement, elle recompacte à chaque fermeture (créant une copie à chaque fois), et elle est **déconseillée pour un back-end situé sur le réseau**, où une interruption peut être dommageable. On la réserve plutôt à un front-end local.

## Compacter le front-end lui-même

Pour compacter le fichier courant — impossible de l'intérieur —, on s'appuie sur un **lanceur externe**. Démarrer Access avec le commutateur `/compact` sur un fichier fermé le compacte sur place puis referme Access ; on peut l'invoquer depuis un script, une tâche planifiée, ou un petit utilitaire de lancement qui compacte le front-end avant de l'ouvrir.

```vba
Shell """C:\Program Files\Microsoft Office\root\Office16\MSACCESS.EXE"" " & _
      """C:\Data\Front.accdb"" /compact", vbHide
```

> ℹ️ Le lancement de processus (`ShellExecute`) est traité au [chapitre 22.2](../22-api-windows-integration-office/02-api-courantes.md), les paramètres de ligne de commande au [chapitre 21.8](../21-deploiement-distribution/08-installation-automatisee.md).

## Précautions et effets de bord

Quelques règles encadrent toute automatisation du compactage. On **sauvegarde** toujours avant l'opération : le compactage réécrit le fichier, et une interruption (coupure, plantage) en cours pourrait l'endommager. On s'assure qu'**aucune session** n'est active sur le fichier visé. On vérifie le résultat (booléen de `CompactRepair`, présence du fichier final) et l'on conserve la sauvegarde tant que le succès n'est pas confirmé.

Un effet de bord mérite d'être connu : le compactage **réinitialise la valeur de départ des champs NuméroAuto** au maximum existant + 1. Si les enregistrements portant les numéros les plus élevés ont été supprimés, la numérotation reprendra à une valeur inférieure à celle attendue — ce qui peut surprendre dans certains contextes métier.

## Points clés à retenir

- Le compactage récupère l'espace perdu, défragmente données et index, et rafraîchit les statistiques de l'optimiseur ; il aide à rester sous la limite des 2 Go.
- On ne peut pas compacter une base ouverte : d'où le compactage d'une autre base, le compactage à la fermeture, ou un lanceur externe.
- `Application.CompactRepair` et `DBEngine.CompactDatabase` compactent un fichier fermé vers un nouveau fichier ; compacter sur place impose un motif d'échange avec sauvegarde.
- Le scénario type est le compactage du back-end par le front-end, en l'absence de toute autre session — opération réservée à une fenêtre de maintenance.
- Le compactage à la fermeture est simple mais déconseillé pour un back-end réseau ; le front-end courant se compacte via le commutateur `/compact` d'un lanceur externe.
- Toujours sauvegarder avant, vérifier le succès, et connaître la réinitialisation de la valeur NuméroAuto.

---


⏭️ [18.8. Optimisation réseau pour les bases partagées (persistance de connexion)](/18-optimisation-performance/08-optimisation-reseau.md)
