🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 2.1. Accès à l'éditeur VBA (Alt+F11) depuis Access

Le chapitre s'ouvre logiquement par la première chose à savoir faire : **atteindre l'éditeur VBA**. Cette section présente toutes les manières d'y accéder depuis Access, la façon de naviguer entre les deux environnements, et propose un premier tour de l'interface. Les fenêtres les plus importantes — explorateur de projets, Propriétés, Exécution immédiate — n'y sont que survolées : elles font l'objet de sections dédiées (2.2 et 2.4).

## L'éditeur VBA, un environnement distinct

Première chose à comprendre : l'éditeur VBA (le **VBE**, *Visual Basic Editor*) n'est pas une fenêtre interne d'Access, mais un **environnement à part entière**, avec sa propre fenêtre, ses propres menus et ses propres barres d'outils. C'est un composant **partagé** par toutes les applications Office : le VBE d'Access est le même outil que celui d'Excel ou de Word, ce qui explique pourquoi l'écriture de code y est familière à qui vient d'une autre application (le langage et l'éditeur sont communs ; seul le modèle objet change, comme l'a montré la section 1.2).

Concrètement, Access et le VBE sont deux fenêtres ouvertes simultanément : on bascule de l'une à l'autre selon que l'on travaille sur les objets (tables, formulaires…) ou sur le code.

## Ouvrir l'éditeur : les différentes voies

Plusieurs chemins mènent au VBE depuis Access.

### Le raccourci Alt+F11 (le plus direct)

<kbd>Alt</kbd>+<kbd>F11</kbd> est de loin la méthode la plus rapide, et la plus utilisée au quotidien. C'est un **basculement** : depuis Access, il ouvre (ou ramène au premier plan) l'éditeur ; depuis l'éditeur, il ramène à Access. On l'utilise en réflexe, des dizaines de fois par session. À ne pas confondre avec <kbd>F11</kbd> seul, qui — autre touche spéciale d'Access — affiche ou masque le **volet de navigation**, et non l'éditeur.

Attention toutefois : ce raccourci dépend de l'option **« Utiliser les touches spéciales Access »** (voir section 1.4). Si elle a été désactivée — typiquement dans une application livrée et verrouillée —, <kbd>Alt</kbd>+<kbd>F11</kbd> reste sans effet. Nous y revenons en fin de section.

### Depuis le ruban

Deux boutons **Visual Basic** ouvrent également l'éditeur :

- ruban **Créer**, groupe **Macros et code** ;
- ruban **Outils de base de données**, groupe **Macro**.

Le ruban **Créer** permet, de surcroît, de créer directement un **Module** ou un **Module de classe**, ce qui ouvre aussitôt l'éditeur sur le nouvel objet.

### En ouvrant un module existant

Dans le **volet de navigation** d'Access, un double-clic sur un module (catégorie *Modules*) l'ouvre directement dans l'éditeur, le curseur prêt dans la fenêtre de code.

### Depuis un formulaire ou un état (le « code derrière »)

C'est l'une des façons les plus fréquentes d'arriver dans le VBE en cours de conception. Lorsqu'on travaille sur un formulaire ou un état en **mode Création** :

- le bouton **Afficher le code** (groupe *Outils* du ruban de création) ouvre le **module de code attaché** à cet objet ;
- pour créer un gestionnaire d'événement, on sélectionne la propriété d'événement voulue dans la **feuille de propriétés** (par exemple *Sur clic* d'un bouton), on clique sur le bouton **« … »** (Générer), puis on choisit **Générateur de code** dans la boîte de dialogue. Access crée alors automatiquement le squelette de la procédure et ouvre l'éditeur dessus.

Cette dernière voie illustre un point important : la feuille de propriétés propose, pour un événement, **trois générateurs** — Générateur d'expression, Générateur de macro et Générateur de code. C'est en choisissant ce dernier que l'on bascule dans le monde du VBA (le choix entre macro et code ayant été discuté en section 1.3).

## Naviguer entre Access et l'éditeur

Une fois l'éditeur ouvert, Access **reste ouvert** en arrière-plan : on ne quitte rien, on bascule. Le moyen le plus simple de revenir reste <kbd>Alt</kbd>+<kbd>F11</kbd>, mais la barre des tâches Windows et le menu Affichage de l'éditeur (commande **Microsoft Access**) permettent aussi de repasser côté application. Cette circulation permanente entre les deux fenêtres — concevoir un formulaire dans Access, écrire son code dans le VBE, tester, revenir — est le rythme normal du développement sous Access.

## Premier tour de l'interface

À la première ouverture, l'éditeur présente une disposition par défaut que l'on peut entièrement personnaliser (fenêtres ancrables, flottantes ou masquées). Voici les éléments principaux à connaître ; les plus importants ont leur propre section.

- **L'explorateur de projets** (raccourci <kbd>Ctrl</kbd>+<kbd>R</kbd>) : l'arborescence des composants de votre projet (modules, modules de classe, modules de formulaire et d'état). Il fait l'objet de la section 2.2.
- **La fenêtre de code** (raccourci <kbd>F7</kbd>) : la zone centrale où l'on écrit et modifie le code. Nous la détaillons juste après.
- **La fenêtre Propriétés** (raccourci <kbd>F4</kbd>) : affiche les propriétés de l'élément sélectionné. Elle est traitée en section 2.4.
- **La fenêtre Exécution immédiate** (raccourci <kbd>Ctrl</kbd>+<kbd>G</kbd>) : pour tester des expressions et recueillir les sorties de `Debug.Print`. Également traitée en section 2.4.
- **L'explorateur d'objets** (raccourci <kbd>F2</kbd>) : permet de parcourir les bibliothèques disponibles, leurs classes et leurs membres — très utile pour explorer le modèle objet et les bibliothèques activées (voir section 2.5).
- **Les fenêtres Variables locales et Espions** : dédiées au débogage, elles sont présentées au chapitre 19.

Au-dessus de ces fenêtres, la **barre de menus** et les **barres d'outils** (Standard, Édition, Débogage…) regroupent les commandes ; on peut afficher ou masquer chaque barre via le menu **Affichage > Barres d'outils**.

## La fenêtre de code et ses listes déroulantes

La fenêtre de code mérite une mention particulière, car elle comporte en haut **deux listes déroulantes** qui facilitent grandement la navigation :

- la liste de **gauche (Objet)** ;
- la liste de **droite (Procédure)**.

Dans un **module standard**, la liste de gauche affiche `(Général)` et celle de droite énumère les procédures et fonctions du module — pratique pour sauter directement à l'une d'elles. Dans un **module de formulaire ou d'état**, la liste de gauche présente le formulaire et ses contrôles, tandis que celle de droite liste les **événements** disponibles pour l'élément choisi : sélectionner un contrôle puis un événement crée (ou rejoint) la procédure correspondante.

```vba
' Fenêtre de code d'un module de formulaire
' Liste de gauche : "btnValider"   —   liste de droite : "Click"
Private Sub btnValider_Click()
    MsgBox "Bouton cliqué"
End Sub
```

Ce couple de listes est le moyen le plus sûr de créer un gestionnaire d'événement avec la **signature exacte** attendue, sans risque d'erreur de nommage.

## Quelques raccourcis essentiels

Au-delà de l'ouverture de l'éditeur, quelques raccourcis reviennent constamment.

| Raccourci | Action |
|---|---|
| <kbd>Alt</kbd>+<kbd>F11</kbd> | Basculer entre Access et l'éditeur VBA |
| <kbd>Ctrl</kbd>+<kbd>R</kbd> | Afficher l'explorateur de projets |
| <kbd>F4</kbd> | Afficher la fenêtre Propriétés |
| <kbd>F7</kbd> | Afficher la fenêtre de code |
| <kbd>Ctrl</kbd>+<kbd>G</kbd> | Afficher la fenêtre Exécution immédiate |
| <kbd>F2</kbd> | Ouvrir l'explorateur d'objets |
| <kbd>F5</kbd> | Exécuter la procédure |
| <kbd>F8</kbd> | Exécuter en pas à pas |

Les touches <kbd>F5</kbd> et <kbd>F8</kbd> relèvent de l'exécution et du débogage (chapitre 19). La liste complète des raccourcis de l'éditeur figure en **annexe I**.

## Quand Alt+F11 ne fonctionne pas

Si <kbd>Alt</kbd>+<kbd>F11</kbd> ne produit aucun effet, la cause la plus fréquente est la désactivation des **touches spéciales Access** (section 1.4), souvent volontaire dans une application finalisée pour empêcher l'utilisateur d'accéder au code et aux objets. C'est précisément l'un des leviers de verrouillage de l'interface décrits en section 17.7. Pour réintervenir sur une telle base en développement, il faut au préalable réactiver ces touches (ou contourner les options de démarrage), ce que nous verrons en temps voulu. Retenez ici le lien de cause à effet : pas de touches spéciales, pas de raccourci vers l'éditeur.

## À retenir

- L'éditeur VBA (**VBE**) est un **environnement distinct** d'Access, partagé par toute la suite Office ; on bascule de l'un à l'autre, sans rien fermer.
- Le moyen le plus rapide de l'ouvrir est <kbd>Alt</kbd>+<kbd>F11</kbd>, un **basculement** qui fonctionne dans les deux sens.
- On y accède aussi par le ruban (**Créer** ou **Outils de base de données** > *Visual Basic*), en ouvrant un module, ou via **Afficher le code** / le **Générateur de code** depuis un formulaire ou un état.
- L'interface s'organise autour de quelques fenêtres clés (explorateur de projets, code, Propriétés, Exécution immédiate), détaillées dans les sections 2.2 et 2.4.
- Les **deux listes déroulantes** de la fenêtre de code (Objet et Procédure) sont le moyen le plus sûr de créer un gestionnaire d'événement correct.
- Si <kbd>Alt</kbd>+<kbd>F11</kbd> reste sans effet, vérifiez l'option **« Utiliser les touches spéciales Access »** (sections 1.4 et 17.7).

---


⏭️ [2.2. Explorateur de projets — modules, formulaires, états](/02-interface-environnement/02-explorateur-projets.md)
