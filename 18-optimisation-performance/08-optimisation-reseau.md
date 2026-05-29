🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 18.8. Optimisation réseau pour les bases partagées (persistance de connexion)

Dès qu'une base Access est partagée, le réseau devient le coût dominant de toute opération. C'est ici que se trouve l'un des leviers de performance les plus spectaculaires de tout le chapitre — la **connexion persistante** —, dont l'effet peut se mesurer en secondes gagnées sur chaque action. Comprendre pourquoi suppose de revenir un instant sur la nature du moteur ACE.

## Le réseau, coût dominant d'une base partagée

Le moteur ACE est un moteur **orienté fichier** : il s'exécute sur le poste de chaque utilisateur, et les données résident dans un fichier back-end sur un partage réseau. Le traitement a donc lieu côté client, et chaque lecture rapatrie des données à travers le réseau. La performance d'une base partagée est par conséquent gouvernée par trois facteurs : le **nombre de fois** où le fichier back-end est ouvert et fermé, le **volume** de données transférées, et la **latence** de la liaison. Optimiser le réseau, c'est agir sur ces trois leviers.

## Prérequis : séparer front-end et back-end

Toute optimisation réseau repose d'abord sur la **séparation** de l'application : un back-end ne contenant que les données partagées, et un front-end — formulaires, requêtes, états, code — installé **localement** sur chaque poste. Les traitements s'exécutent alors en local, et seules les données transitent par le réseau, là où un fichier monolithique partagé ferait circuler bien davantage. Cette architecture est le socle sur lequel tout le reste s'appuie.

> ℹ️ L'architecture front-end / back-end est détaillée au [chapitre 15.1](../15-multi-utilisateurs/01-architecture-front-end-back-end.md).

## Le levier majeur : la connexion persistante

### Le problème : ouvrir et fermer le back-end sans cesse

Lorsqu'aucun objet ne maintient le back-end ouvert, Access l'ouvre et le ferme **à répétition** — à chaque requête, chaque Recordset, chaque recherche. Or ouvrir un fichier sur le réseau n'est pas anodin : négociation de verrouillage, lecture de l'en-tête, création puis suppression du fichier de verrouillage… Répétée des centaines de fois, cette danse d'ouverture/fermeture coûte un temps considérable.

### La solution : garder le back-end ouvert toute la session

Le remède consiste à maintenir **une connexion ouverte** vers le back-end pendant toute la durée de la session du front-end. Le fichier n'est alors ouvert qu'une seule fois ; toutes les opérations suivantes le réutilisent. La technique classique : ouvrir, au démarrage de l'application, un Recordset léger sur une table liée du back-end, et le conserver dans une variable de module jusqu'à la fermeture.

```vba
' === modConnexion (module standard) ===
Private mrsBackEnd As DAO.Recordset

Public Sub OuvrirConnexionPersistante()
    If mrsBackEnd Is Nothing Then
        ' Recordset léger sur une table liée : maintient le back-end ouvert
        Set mrsBackEnd = CurrentDb.OpenRecordset( _
            "SELECT TOP 1 * FROM Parametres", dbOpenForwardOnly)
    End If
End Sub

Public Sub FermerConnexionPersistante()
    On Error Resume Next
    If Not mrsBackEnd Is Nothing Then
        mrsBackEnd.Close
        Set mrsBackEnd = Nothing
    End If
End Sub
```

On appelle `OuvrirConnexionPersistante` au lancement (formulaire de démarrage ou routine d'autoexec) et `FermerConnexionPersistante` à la fermeture. La connexion reste en mode **partagé** (le mode par défaut), afin que les autres utilisateurs puissent se connecter normalement ; elle stabilise au passage le fichier de verrouillage. Une variante consiste à garder ouvert, masqué, un formulaire lié à une table du back-end : l'effet est identique.

> ℹ️ Les modes de partage sont traités au [chapitre 15.2](../15-multi-utilisateurs/02-modes-partage.md), les tables liées par code au [chapitre 15.9](../15-multi-utilisateurs/09-tables-liees-par-code.md).

## Désactiver les sous-feuilles de données

Un autre coupable réseau, souvent ignoré : les **sous-feuilles de données** (subdatasheets). Quand la propriété `SubdatasheetName` d'une table vaut `[Auto]`, Access ouvre les tables liées pour préparer l'affichage déroulant à chaque ouverture de la table — générant un trafic réseau superflu. Régler cette propriété à `[None]` sur toutes les tables supprime cette surcharge.

```vba
Public Sub DesactiverSousFeuilles(ByVal db As DAO.Database)
    Dim td As DAO.TableDef, prp As DAO.Property
    For Each td In db.TableDefs
        If (td.Attributes And dbSystemObject) = 0 Then       ' ignorer les tables système
            On Error Resume Next
            td.Properties("SubdatasheetName") = "[None]"
            If Err.Number = 3270 Then                        ' propriété absente : la créer
                Err.Clear
                Set prp = td.CreateProperty("SubdatasheetName", dbText, "[None]")
                td.Properties.Append prp
            End If
            On Error GoTo 0
        End If
    Next td
End Sub
```

Ce réglage se pose sur les tables du **back-end** (ou se fixe une fois pour toutes à la conception).

> ℹ️ La collection `TableDefs` est traitée au [chapitre 12.4](../12-querydefs-tabledefs/04-collection-tabledefs.md), les propriétés personnalisées des objets au [chapitre 12.8](../12-querydefs-tabledefs/08-proprietes-personnalisees.md).

## Transférer moins de données

Sur le réseau, le coût de la sur-récupération est amplifié. Les principes des [sections 18.2](02-optimisation-recordsets.md) et [18.5](05-optimisation-formulaires.md) prennent ici tout leur poids : ne demander que les colonnes et les lignes utiles, filtrer **à la source** (clause `WHERE`) plutôt que côté client, et éviter de charger des tables entières dans des formulaires ou des listes déroulantes. Chaque ligne et chaque colonne épargnées sont autant d'octets qui ne traversent pas le réseau.

## Traiter en local quand c'est possible

Pour un traitement en plusieurs étapes manipulant des résultats intermédiaires, on gagne à effectuer le travail dans des **tables temporaires locales** (dans le front-end ou une base de travail locale), puis à ne renvoyer que le résultat final vers le back-end. On évite ainsi de multiplier les allers-retours réseau pendant les calculs intermédiaires. Un Recordset ADO déconnecté peut jouer un rôle comparable.

> ℹ️ Voir les Recordsets déconnectés au [chapitre 10.7](../10-ado-access/07-recordsets-deconnectes.md).

## Ce que le réseau ne pardonne pas

Le moteur ACE suppose une liaison **fiable et à faible latence**, c'est-à-dire un réseau local câblé. Un back-end Access utilisé à travers le Wi-Fi, un VPN ou un lien WAN est réputé lent et instable : la latence et les pertes de paquets dégradent les performances et exposent à un risque de corruption du fichier. Lorsque ces conditions sont imposées — postes distants, multi-sites, forte concurrence —, la bonne réponse n'est plus d'optimiser le réseau mais de migrer les données vers un moteur serveur, où le traitement s'effectue côté serveur et où seul le résultat circule.

> ℹ️ Les limites du moteur ACE en environnement concurrent sont exposées au [chapitre 15.10](../15-multi-utilisateurs/10-limites-moteur-ace.md) ; la migration vers SQL Server au [chapitre 23](../23-migration-interoperabilite/README.md). Le déploiement réseau et les chemins UNC sont traités au [chapitre 21.5](../21-deploiement-distribution/05-deploiement-reseau-chemins-unc.md).

## Points clés à retenir

- Pour une base partagée, le réseau domine : le coût dépend du nombre d'ouvertures du fichier, du volume transféré et de la latence.
- La séparation front-end / back-end, avec un front-end local, est le prérequis de toute optimisation réseau.
- La connexion persistante — garder un Recordset (ou un formulaire masqué) ouvert sur le back-end toute la session — évite les ouvertures/fermetures répétées : c'est le gain le plus net.
- Régler `SubdatasheetName` à `[None]` sur les tables supprime un trafic réseau inutile.
- Transférer moins (colonnes, lignes, filtrage à la source) compte d'autant plus sur le réseau ; traiter les étapes intermédiaires en local réduit les allers-retours.
- Le moteur ACE exige un LAN câblé fiable ; pour le distant, le VPN ou la forte concurrence, la solution est la migration vers un serveur.

---


⏭️ [18.9. Limites de taille (2 Go) — stratégies de contournement](/18-optimisation-performance/09-limites-taille-2go.md)
