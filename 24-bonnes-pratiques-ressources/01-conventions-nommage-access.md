🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 24.1. Conventions de nommage spécifiques Access (préfixes Leszynski/Reddick)

Un nom bien choisi est la première forme de documentation. Lorsqu'on lit `cboPays`, on sait sans rien vérifier qu'il s'agit d'une liste déroulante (combo box) servant à choisir un pays ; lorsqu'on lit `tblClients`, qu'il s'agit d'une table. Cette lisibilité immédiate, multipliée par les centaines d'objets et de variables d'une application, fait gagner un temps considérable en lecture, en maintenance et en travail d'équipe. C'est l'objet des conventions de nommage, et dans le monde Access, deux noms reviennent invariablement : **Leszynski** et **Reddick**. Mais au-delà du confort général, le nommage répond dans Access à des contraintes bien spécifiques, que cette section met en lumière.

## D'où viennent ces conventions

Stan Leszynski et Greg Reddick ont élaboré, dans les années 1990, une convention de nommage commune, dont chacun a ensuite décliné sa propre variante. La **Leszynski Naming Convention** (LNC) est centrée sur Access, tandis que les **Reddick VBA Naming Conventions** (RVBA) couvrent plus largement VBA. Toutes deux partagent la même philosophie et sont devenues le standard de fait de la communauté Access. Il s'agit d'une forme de *notation hongroise* adaptée : on préfixe le nom d'un objet par une courte balise indiquant sa nature.

## L'anatomie d'un nom

Un nom complet suit une structure régulière : un ou plusieurs **préfixes** optionnels, une **balise** (tag), un **nom de base**, et éventuellement un **qualificatif**.

```
[préfixe(s)] [balise] NomDeBase [qualificatif]
     g          str    NomUtilisateur            ->  gstrNomUtilisateur
```

La **balise** est le cœur du système : une abréviation courte, en minuscules, qui désigne le type ou la nature de l'élément (`tbl`, `frm`, `txt`, `str`…). Le **nom de base** décrit ce que l'élément représente, écrit en PascalCase et de façon explicite (`Clients`, `NomUtilisateur`, et non `cli` ou `nu`). Les **préfixes**, réservés surtout aux variables, indiquent la portée ou la durée de vie. Le **qualificatif** précise au besoin (`_Max`, `Ancien`, `Nouveau`). Ainsi `txtNomClient` est une zone de texte affichant le nom du client, et `gstrNomUtilisateur` une variable globale de type chaîne contenant le nom de l'utilisateur.

## Les balises des objets de base de données

| Objet | Balise | Exemple |
|---|---|---|
| Table | `tbl` | `tblClients` |
| Requête | `qry` | `qryCommandesEnCours` |
| Formulaire | `frm` | `frmSaisieCommande` |
| Sous-formulaire | `fsub` | `fsubLignesCommande` |
| État | `rpt` | `rptFactureMensuelle` |
| Macro | `mcr` | `mcrDemarrage` |
| Module standard | `mod` (ou `bas`) | `modUtilitaires` |
| Module de classe | `cls` | `clsClient` |

Certaines variantes affinent encore les requêtes selon leur type (`qrySel` pour une requête Sélection, `qryApp` pour Ajout, `qryUpd` pour Mise à jour, `qryDel` pour Suppression), mais beaucoup d'équipes s'en tiennent à `qry`. La liste exhaustive figure à l'annexe G.

## Les balises des contrôles

| Contrôle | Balise | Exemple |
|---|---|---|
| Étiquette (label) | `lbl` | `lblTitre` |
| Zone de texte | `txt` | `txtNomClient` |
| Liste déroulante (combo) | `cbo` | `cboPays` |
| Zone de liste | `lst` | `lstProduits` |
| Bouton de commande | `cmd` | `cmdEnregistrer` |
| Case à cocher | `chk` | `chkActif` |
| Bouton d'option | `opt` | `optHomme` |
| Groupe d'options | `grp` (ou `ogp`) | `grpSexe` |
| Bouton bascule | `tgl` | `tglAffichage` |
| Image | `img` | `imgLogo` |
| Contrôle onglet | `tab` | `tabFiche` |
| Page d'onglet | `pge` | `pgeCoordonnees` |
| Sous-formulaire (contrôle) | `sub` | `subLignes` |

## Les balises des variables et des objets de données

Pour les variables, la balise reflète le **type de données**, et le préfixe la **portée**.

| Type de données | Balise | | Portée | Préfixe |
|---|---|---|---|---|
| String | `str` | | Globale / publique | `g` |
| Integer | `int` | | Niveau module (privée) | `m` |
| Long | `lng` | | Statique | `s` |
| Boolean | `bln` | | Locale | *(aucun)* |
| Currency | `cur` | | | |
| Double | `dbl` | | | |
| Single | `sng` | | | |
| Date | `dat` | | | |
| Variant | `var` | | | |
| Object | `obj` | | | |

Les objets DAO et ADO ont eux aussi leurs balises consacrées : `db` pour une Database, `rst` pour un Recordset, `qdf` pour un QueryDef, `tdf` pour un TableDef, `fld` pour un Field, `idx` pour un Index. Là encore, la référence complète est en annexe G. À noter une légère divergence entre conventions : la LNC utilise parfois `f` (pour *flag*) là où Reddick préfère `bln` pour un booléen ; l'essentiel est de choisir l'une et de s'y tenir.

## Pourquoi ces préfixes sont particulièrement utiles dans Access

Ici se trouve la véritable raison d'être de cette section. Si la notation hongroise est discutée ailleurs, elle répond dans Access à trois problèmes concrets qui, eux, ne sont pas discutables.

### Le piège du contrôle qui masque le champ

C'est l'argument le plus fort. Lorsqu'on dépose un champ sur un formulaire, Access nomme par défaut le contrôle **du même nom que le champ**. Or si un contrôle calculé, une expression ou une autre référence utilise ce nom, Access ne sait plus s'il s'agit du champ ou du contrôle : il en résulte une **référence circulaire** et l'affichage de `#Erreur` ou `#Nom?`.

```
' Champ de la table : Total
' MAUVAIS : le contrôle se nomme aussi "Total"
'   -> référence circulaire, #Erreur
' BON : le contrôle est préfixé et distinct du champ
'   txtTotal  (avec ControlSource = Total)
```

Préfixer systématiquement les contrôles (`txt`, `cbo`, `cmd`…) élimine purement et simplement cette catégorie de bugs, parmi les plus déroutants pour les débutants. À lui seul, ce point justifie la convention.

### Les collisions dans l'espace de noms

Access ne permet pas qu'une table et une requête portent le même nom, et mélanger des noms identiques entre objets de natures différentes conduit à des conflits et à des références ambiguës. Préfixer (`tblClients`, `qryClients`, `frmClients`) sépare proprement les espaces de noms et supprime ces ambiguïtés.

### L'organisation du volet de navigation

Comme le volet de navigation trie les objets par ordre alphabétique, les préfixes **regroupent naturellement** les objets de même nature : toutes les tables sous `tbl…`, toutes les requêtes sous `qry…`, tous les formulaires sous `frm…`. Sur une application comportant des dizaines d'objets, ce regroupement automatique facilite grandement le repérage.

## Le débat moderne et la juste mesure

Il faut être honnête : la notation hongroise a perdu de sa popularité dans le développement logiciel en général. Les langages et environnements modernes, avec leur IntelliSense, rendent les préfixes de type souvent redondants, et beaucoup de développeurs les jugent encombrants. Cette critique est recevable… ailleurs. Dans Access, les trois raisons exposées ci-dessus — masquage des champs, collisions de noms, organisation du volet — restent parfaitement valables et propres à l'outil.

La position raisonnable est donc nuancée. Préfixer les **objets de base de données** et les **contrôles** repose sur des justifications solides et spécifiques à Access : c'est fortement recommandé. Préfixer les **variables** par leur type est plus discutable, mais demeure une pratique répandue et cohérente en VBA. Dans tous les cas, la règle cardinale n'est pas le dogme, mais la **cohérence** : une convention appliquée à moitié vaut moins qu'une convention simple appliquée partout.

## Recommandations pratiques

En pratique, quelques principes suffisent. **Choisir une convention** — LNC, Reddick, ou un sous-ensemble simplifié — et la **documenter** pour toute l'équipe, afin que chacun nomme de la même façon. **Préfixer les objets et les contrôles** sans exception. Pour les variables, faire précéder la balise de type d'un préfixe de portée (`g`, `m`). Donner des **noms de base explicites en PascalCase**, sans abréviations cryptiques : `txtNomClient` et non `txtNC`. Et ne pas tomber dans l'excès inverse, en empilant des balises au point de rendre les noms illisibles. La référence exhaustive des préfixes se trouve à l'**annexe G**, à garder sous la main.

## En résumé

Les conventions Leszynski et Reddick structurent les noms en *balise + nom de base*, éventuellement précédés d'un préfixe de portée pour les variables. Au-delà du confort de lecture commun à tout langage, elles répondent dans Access à des besoins spécifiques et bien réels : éviter le piège du contrôle qui masque son champ, prévenir les collisions de noms, et organiser le volet de navigation. Préfixer objets et contrôles est donc fortement justifié ; pour les variables, l'usage reste répandu et la cohérence prime sur tout. Une fois ce vocabulaire commun adopté, la cohérence du nommage prépare le terrain à la question plus large de la structure d'ensemble de l'application — son architecture, objet de la section suivante.

⏭️ [24.2. Architecture applicative recommandée pour Access](/24-bonnes-pratiques-ressources/02-architecture-applicative.md)
