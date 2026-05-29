🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 15.2. Modes de partage de la base de données

La manière dont une base Access est **ouverte** détermine qui d'autre peut y accéder en même temps. C'est le point de départ concret du multi-utilisateur : avant de parler de stratégies de verrouillage, il faut comprendre les modes d'ouverture — exclusif, partagé, lecture seule — et savoir comment les contrôler, depuis l'interface, en ligne de commande ou par code. Cette section précise aussi les réglages globaux qui encadrent le partage.

## Les deux modes fondamentaux : exclusif et partagé

Une base Access s'ouvre selon l'un de deux modes principaux.

En mode **exclusif**, un seul utilisateur peut ouvrir la base à la fois. Tant qu'elle est ouverte exclusivement, aucune autre session ne peut s'y connecter — toute tentative se solde par un message d'erreur. Ce mode est réservé aux opérations d'administration qui exigent un accès sans concurrence : définir un mot de passe de base, compacter, ou modifier certaines structures partagées.

En mode **partagé**, plusieurs utilisateurs peuvent ouvrir la base simultanément. C'est le mode requis pour le fonctionnement multi-utilisateur, et c'est le mode par défaut. Dans une architecture front-end / back-end ([section 15.1](01-architecture-front-end-back-end.md)), c'est ce mode qui permet à plusieurs front-ends d'accéder ensemble au back-end via leurs tables liées.

## Le mode lecture seule

En complément, une base peut être ouverte en **lecture seule** : les données peuvent être consultées mais non modifiées. Ce mode est utile pour de la consultation pure, lorsqu'on veut écarter tout risque d'écriture accidentelle, ou lorsque le support lui-même n'autorise pas l'écriture. Il se combine avec les modes précédents (lecture seule partagée, lecture seule exclusive).

## Définir le mode d'ouverture par défaut

Le mode appliqué par défaut à l'ouverture des bases est un réglage de l'installation Access, accessible par Fichier → Options → Paramètres client → Avancé → « Mode d'ouverture par défaut » (Partagé ou Exclusif).

Ce réglage est également lisible et modifiable par code, via les méthodes `GetOption` et `SetOption` de l'objet `Application`, qui donnent accès à toutes les options de la boîte de dialogue Options. L'option concernée porte le nom `"Default Open Mode for Databases"`, dont la valeur `0` correspond à Partagé et `1` à Exclusif.

```vba
' Lire le mode d'ouverture par défaut (0 = Partagé, 1 = Exclusif)
Debug.Print "Mode : " & Application.GetOption("Default Open Mode for Databases")

' Forcer le mode Partagé
Application.SetOption "Default Open Mode for Databases", 0
```

Pour une option de type liste, `GetOption` renvoie la **position** (commençant à zéro) du choix sélectionné, et `SetOption` attend cette position numérique. Deux précautions accompagnent ces méthodes : il n'est pas nécessaire d'indiquer l'onglet où figure l'option, et — point souvent négligé — **les noms d'options doivent toujours être fournis en anglais**, même sur une version localisée d'Access, sous peine de ne pas être reconnus.

## Forcer le mode à l'ouverture

Au-delà du défaut, plusieurs moyens permettent d'imposer un mode pour une ouverture donnée.

**Depuis l'interface**, le bouton Ouvrir propose un menu déroulant offrant les variantes : Ouvrir, Ouvrir en lecture seule, Ouvrir en mode exclusif, Ouvrir en mode exclusif et en lecture seule.

**En ligne de commande**, des commutateurs permettent de contrôler le mode dans le raccourci qui lance l'application — particulièrement utile pour le déploiement d'un front-end : le commutateur `/excl` ouvre en mode exclusif, et `/ro` en lecture seule. On les place à la suite du chemin du fichier dans la cible du raccourci :

```
"C:\Applications\MonAppli.accdb" /excl
"C:\Applications\MonAppli.accdb" /ro
```

**Par code**, lorsqu'on ouvre une *autre* base depuis VBA (par exemple un back-end), la méthode `OpenDatabase` du `DBEngine` accepte le mode en paramètre. Sa signature est `OpenDatabase(Nom, Exclusif, LectureSeule, Connect)` : le deuxième argument booléen vaut `True` pour exclusif, `False` pour partagé.

```vba
Dim db As DAO.Database
' Ouvrir un back-end en mode PARTAGÉ (Exclusif = False), en lecture/écriture
Set db = DBEngine.OpenDatabase("\\Serveur\Partage\MaBase_be.accdb", False, False)
' ... traitement ...
db.Close
Set db = Nothing
```

## Permissions sur le dossier partagé

Un prérequis du partage est souvent sous-estimé : pour qu'une base soit ouverte en mode partagé en écriture, chaque utilisateur doit disposer des droits de **lecture, écriture, création et suppression** sur le **dossier** contenant la base — et pas seulement sur le fichier lui-même. La raison est que le moteur doit pouvoir y créer, puis y supprimer, le fichier de verrouillage. Un dossier en lecture seule empêche le partage en écriture, voire l'ouverture si le fichier de verrouillage ne peut être créé. C'est une cause classique de blocage au déploiement.

## Le fichier de verrouillage

Lorsqu'une base est ouverte en mode partagé, Access crée à côté d'elle un **fichier de verrouillage** portant le même nom et l'extension `.laccdb` (pour une base `.accdb`) ou `.ldb` (pour une base `.mdb`). Ce fichier ne contient aucune donnée : il sert à recenser les sessions ouvertes et à coordonner le partage. Il est normalement supprimé automatiquement à la fermeture de la dernière session, même s'il arrive qu'il subsiste dans certains cas. Le suivi des sessions à partir de ce fichier est traité à la [section 15.8](08-gestion-sessions-utilisateurs.md).

## Les réglages de verrouillage associés

Deux autres réglages globaux, situés au même endroit (Paramètres client → Avancé), conditionnent le comportement concurrent. Ils sont présentés ici en tant que **valeurs par défaut** ; le choix de stratégie proprement dit est développé à la [section 15.3](03-strategies-verrouillage.md), et le contrôle au niveau des formulaires à la [section 15.6](06-recordlocks-formulaires.md).

Le **verrouillage des enregistrements par défaut** (`"Default Record Locking"`) offre trois choix : « Aucun verrou » (valeur `0`, verrouillage optimiste), « Tous les enregistrements » (`1`) et « Enregistrement modifié » (`2`, verrouillage pessimiste).

Le **verrouillage au niveau enregistrement** (`"Use Row Level Locking"`) est un commutateur qui détermine la granularité du verrouillage : activé, le moteur verrouille l'enregistrement ; désactivé, il verrouille la page entière qui le contient (et donc potentiellement des enregistrements voisins).

```vba
' Réglages typiques d'un déploiement multi-utilisateur
Application.SetOption "Default Record Locking", 0     ' Aucun verrou (optimiste)
Application.SetOption "Use Row Level Locking", True   ' Verrouillage par enregistrement
```

## Réglages du moteur

À un niveau plus bas, certains réglages du moteur ACE lui-même — distincts des options de l'interface — influencent le comportement en environnement partagé. Ils se pilotent par la méthode `SetOption` du `DBEngine`, avec des constantes telles que `dbMaxLocksPerFile` (nombre maximal de verrous suivis simultanément), `dbLockRetry` ou `dbPageTimeout`. Ces paramètres relèvent davantage de l'optimisation et sont approfondis au [chapitre 18](/18-optimisation-performance/README.md) ; il suffit ici de savoir qu'ils existent et qu'ils agissent sur le moteur, là où `Application.SetOption` agit sur les options de l'application.

## Interaction avec l'architecture front-end / back-end

Dans le schéma recommandé, chaque front-end s'ouvre en mode partagé, et le back-end est sollicité en mode partagé au travers des tables liées. Tant qu'aucune session ne détient le back-end en exclusif, tous les front-ends y accèdent simultanément.

La contrepartie concerne la **maintenance du back-end** : compacter le fichier de données, en modifier la structure ou y poser un mot de passe exige un accès **exclusif** au back-end — ce qui suppose que tous les front-ends en soient déconnectés au préalable. Prévoir une fenêtre de maintenance, ou un mécanisme pour inviter les utilisateurs à se déconnecter, fait partie de la conception d'une application partagée.

## Points de vigilance

- **Partagé pour le multi-utilisateur, exclusif pour l'administration.** L'exclusif bloque tout le monde ; on ne l'emploie que pour les opérations qui l'exigent réellement.
- **Droits complets sur le dossier**, pas seulement sur le fichier : le moteur doit pouvoir créer et supprimer le fichier de verrouillage.
- **Noms d'options en anglais** dans `GetOption`/`SetOption`, quelle que soit la langue d'Access.
- **`GetOption` renvoie une position** (indexée à zéro) pour les options de type liste, pas un libellé.
- **Maintenance du back-end = accès exclusif** : déconnecter tous les front-ends avant compactage ou modification de structure.
- **Le commutateur `/excl` dans un raccourci de front-end** est rarement souhaitable en multi-utilisateur : il empêcherait les autres d'ouvrir leur propre session si la base visée était partagée.

## En résumé

Une base Access s'ouvre en mode **exclusif** (un seul utilisateur, réservé à l'administration) ou **partagé** (plusieurs utilisateurs, mode requis pour le multi-utilisateur), éventuellement en **lecture seule**. Le mode par défaut se règle dans les Paramètres client ou par code via `Application.SetOption "Default Open Mode for Databases"` (0 = Partagé, 1 = Exclusif), et peut être forcé à l'ouverture par le menu Ouvrir, par les commutateurs en ligne de commande `/excl` et `/ro`, ou par le paramètre d'exclusivité de `DBEngine.OpenDatabase`. Le partage exige des droits complets sur le **dossier** (pour le fichier de verrouillage `.laccdb`/`.ldb`), et s'accompagne de réglages de verrouillage par défaut — verrouillage des enregistrements et granularité par enregistrement — dont la stratégie est détaillée dans la section suivante.

La section suivante approfondit justement ce choix : les [stratégies de verrouillage (optimiste vs pessimiste)](03-strategies-verrouillage.md).

⏭️ [15.3. Stratégies de verrouillage (Optimistic vs Pessimistic)](/15-multi-utilisateurs/03-strategies-verrouillage.md)
