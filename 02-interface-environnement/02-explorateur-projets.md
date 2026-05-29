🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 2.2. Explorateur de projets — modules, formulaires, états

La section précédente a montré comment ouvrir l'éditeur et a survolé ses fenêtres. Parmi elles, l'**explorateur de projets** mérite une attention particulière : c'est la **carte** de tout le code de votre base. On l'affiche par <kbd>Ctrl</kbd>+<kbd>R</kbd> (ou via le menu **Affichage > Explorateur de projets**). Cette section détaille ce qu'il contient, comment s'y repérer et comment y agir.

## Qu'est-ce qu'un projet VBA dans Access ?

À chaque base de données Access correspond **un projet VBA**, qui regroupe tout son code. Par défaut, ce projet porte le nom du fichier de base, sans son extension, et l'on peut le renommer — nous y reviendrons. L'explorateur de projets affiche ce projet et l'ensemble de ses composants de code sous forme d'**arborescence**.

À la différence d'Excel, où l'on jongle souvent avec plusieurs classeurs ouverts (chacun avec son projet), un complément personnel (`Personal.xlsb`) et des macros complémentaires, on travaille généralement, sous Access, avec **un seul projet à la fois** — celui de la base ouverte. C'est plus simple, mais cela ne dispense pas de l'organiser soigneusement dès que l'application grossit.

## Ce que contient l'explorateur de projets

L'arborescence d'un projet Access s'organise en quelques dossiers, qui n'apparaissent que s'ils contiennent quelque chose :

```
NomDeLaBase (projet)
├── Microsoft Access Class Objects
│   ├── Form_frmClients
│   ├── Form_frmCommandes
│   └── Report_rptFactures
├── Modules
│   ├── modUtilitaires
│   └── modAccesDonnees
└── Class Modules
    └── clsClient
```

- **Microsoft Access Class Objects** : les modules de code attachés aux **formulaires** (`Form_…`) et aux **états** (`Report_…`).
- **Modules** : les **modules standard**, qui contiennent les procédures et fonctions générales.
- **Class Modules** : les **modules de classe** personnalisés, lorsque l'application en définit (chapitre 16).

Les différents types de modules et leurs rôles sont précisés à la section suivante (2.3).

> ⚠️ **Point essentiel, source de confusion fréquente.** L'explorateur de projets ne montre **que le code**. Les **tables**, **requêtes** et **macros** n'y figurent **pas** : ce sont des objets Access, qui restent dans le **volet de navigation** d'Access (section 1.5). Cette répartition prolonge le modèle « deux environnements » décrit en section 2.1 : les objets côté Access, le code côté éditeur.

## Modules de formulaire et d'état : le « code derrière »

Les nœuds `Form_frmClients` ou `Report_rptFactures` représentent le **code attaché** à un formulaire ou à un état — ce que l'on appelle le « code derrière » (*code-behind*). C'est là que vivent les gestionnaires d'événements de l'objet (clic, ouverture, mise à jour…).

Il faut bien distinguer deux choses : le **module de code**, visible dans l'explorateur de projets côté éditeur, et la **conception visuelle** du formulaire ou de l'état, qui se trouve dans Access. Les deux sont liés mais résident dans des environnements différents — d'où le bouton **Afficher l'objet** évoqué plus loin, qui fait le pont de l'un à l'autre.

Le nom de ces modules est **imposé** par Access : il dérive du nom de l'objet, préfixé de `Form_` ou `Report_`. On ne le choisit donc pas, contrairement aux modules standard et de classe, que l'on nomme librement.

## La propriété « Contient un module » (HasModule)

Un détail important explique pourquoi certains formulaires ou états **n'apparaissent pas** dans *Microsoft Access Class Objects* : tout simplement, ils n'ont **pas de module de code**. Les formulaires et états possèdent une propriété **« Contient un module »** (`HasModule`) :

- tant qu'elle vaut **Non**, aucun module n'est associé, et l'objet est absent de l'explorateur de projets ;
- dès que l'on crée un gestionnaire d'événement (ou que l'on passe la propriété à **Oui**), Access crée le module et l'objet apparaît.

Cette propriété a une incidence sur les **performances** : un formulaire sans module (dit « léger ») se charge plus vite et alourdit moins la base. C'est l'une des optimisations abordées en section 18.5. Retenez ici la règle de lecture : un formulaire absent de l'explorateur n'est pas une anomalie — il n'a tout simplement pas de code.

## Naviguer et agir depuis l'explorateur

L'explorateur de projets est le principal outil de navigation dans le code. En pratique :

- un **double-clic** sur un module ouvre sa fenêtre de code ;
- trois boutons figurent en haut de la fenêtre :
  - **Afficher le code** : ouvre le code de l'élément sélectionné ;
  - **Afficher l'objet** : pour un module de formulaire ou d'état, **rebascule vers Access** pour montrer la conception de l'objet correspondant ;
  - **Basculer les dossiers** : affiche ou masque le regroupement par dossiers (vue arborescente *vs* liste à plat).

Le bouton **Afficher l'objet** matérialise concrètement le pont entre code et conception : on passe ainsi du gestionnaire d'événement, dans l'éditeur, au formulaire qu'il anime, dans Access.

## Renommer le projet

Le nom du projet — celui qui coiffe l'arborescence — peut être modifié via **Outils > Propriétés de [nom du projet]…**, onglet **Général**, champ **Nom du projet**. Ce n'est pas qu'un détail cosmétique : le nom du projet **compte** dès lors que l'on référence d'autres bases. Deux projets portant le même nom ne peuvent pas se référencer mutuellement, et comme le projet hérite du nom du fichier, une base au nom peu distinctif (« Database.accdb »…) produit un projet tout aussi générique, source de conflits. Donner à chaque projet un nom **explicite et unique** est donc une bonne pratique, sur laquelle nous reviendrons avec les références (section 2.5).

## Plusieurs projets dans l'explorateur

L'explorateur affiche le plus souvent un seul projet, mais il peut en contenir **plusieurs**. C'est le cas lorsque la base en **référence une autre** (par exemple une base d'utilitaires partagée ou un complément) : le projet référencé apparaît alors comme un nœud distinct, généralement en lecture seule. Cette situation reste minoritaire dans une application Access classique ; nous la mentionnons pour que l'apparition d'un second projet ne surprenne pas.

> ℹ️ À ne pas confondre : les **bibliothèques** COM activées (DAO, ADO, Scripting…) **ne s'affichent pas** comme des projets dans l'explorateur. Elles se gèrent ailleurs, via **Outils > Références**, et s'explorent dans l'**explorateur d'objets** (sections 2.5 et 2.1).

## L'explorateur de projets et la fenêtre Propriétés

Enfin, l'explorateur travaille de pair avec la **fenêtre Propriétés** (<kbd>F4</kbd>) : sélectionner un module y affiche ses propriétés. Pour un module standard, son **nom** (`Name`) en est l'unique propriété éditable — c'est d'ailleurs le moyen le plus simple de renommer un module ; un module de classe en expose une seconde, `Instancing`. La fenêtre Propriétés est détaillée à la section 2.4.

## À retenir

- L'**explorateur de projets** (<kbd>Ctrl</kbd>+<kbd>R</kbd>) est la **carte du code** de la base : à chaque base correspond **un projet VBA**.
- Il organise le code en trois dossiers : **Microsoft Access Class Objects** (modules de formulaire et d'état), **Modules** (standard) et **Class Modules** (classes personnalisées).
- **Seul le code y figure** : tables, requêtes et macros restent dans le volet de navigation d'Access.
- Un formulaire ou un état n'apparaît que s'il **possède un module** (`HasModule` = Oui) ; les formulaires « légers » sans code en sont absents et se chargent plus vite (section 18.5).
- Le bouton **Afficher l'objet** fait le pont entre le code (éditeur) et la conception (Access) ; **Afficher le code** et **Basculer les dossiers** complètent la navigation.
- Donnez au projet un **nom explicite et unique** (Outils > Propriétés), surtout si vous prévoyez des références entre bases (section 2.5).

---


⏭️ [2.3. Types de modules : Standard, Classe, Form, Report](/02-interface-environnement/03-types-modules.md)
