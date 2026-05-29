🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 2.4. Fenêtre Propriétés et fenêtre Exécution immédiate

Après l'explorateur de projets et les types de modules, voici deux fenêtres de l'éditeur que tout développeur côtoie — mais à des degrés très inégaux. La **fenêtre Propriétés** rend, dans Access, des services limités ; la **fenêtre Exécution immédiate**, en revanche, est l'un des outils les plus utiles au quotidien. Cette section présente les deux, en insistant sur la seconde, et lève au passage une confusion courante autour du mot « propriétés ».

## La fenêtre Propriétés

### Dans l'éditeur : un usage limité

La fenêtre Propriétés de l'éditeur s'affiche par <kbd>F4</kbd> (ou via **Affichage > Fenêtre Propriétés**). Elle montre les propriétés de l'élément sélectionné dans l'explorateur de projets. Elle propose deux onglets, **Alphabétique** et **Par catégorie**, surtout utiles dans d'autres applications Office.

Sous Access, son intérêt est **restreint** : pour un module standard, la seule propriété disponible est **`(Name)`**, le nom du module ; un module de classe en expose une seconde, **`Instancing`** (réglage avancé, laissé sur *Private* dans la quasi-totalité des cas). En pratique, on ne s'en sert donc guère que pour **renommer un module** — ce qui reste, il est vrai, le moyen le plus simple de le faire.

Cette pauvreté n'est pas un défaut : elle s'explique par la nature d'Access. Comme on l'a vu en sections 1.2 et 2.2, les **formulaires et états ne se conçoivent pas dans l'éditeur** (contrairement aux UserForms d'Excel), mais dans Access. Leurs nombreuses propriétés n'ont donc pas leur place dans la fenêtre Propriétés du VBE.

### Ne pas confondre avec la feuille de propriétés d'Access

C'est ici qu'une **confusion fréquente** doit être levée. Il existe en réalité **deux fenêtres distinctes**, et toutes deux s'ouvrent avec <kbd>F4</kbd> :

- dans **l'éditeur VBA** : la **fenêtre Propriétés** décrite ci-dessus (limitée, pour l'essentiel, au nom des modules) ;
- dans **Access**, sur un formulaire ou un état en mode Création : la **feuille de propriétés** (*Property Sheet*), bien plus riche, qui expose **toutes** les propriétés de l'objet et de ses contrôles (source de données, format, événements, mise en forme…).

Même touche, mais deux fenêtres, dans deux environnements différents. Lorsqu'on parle des « propriétés d'un formulaire » ou d'un contrôle, c'est presque toujours de la **feuille de propriétés d'Access** qu'il s'agit. Celle-ci est étudiée avec les formulaires et les contrôles aux **chapitres 6 et 8**, et les propriétés essentielles sont recensées en **annexe B**.

## La fenêtre Exécution immédiate

Changement de registre : la fenêtre Exécution immédiate est un outil **central**. On l'affiche par <kbd>Ctrl</kbd>+<kbd>G</kbd> (ou via **Affichage > Fenêtre Exécution immédiate**). C'est à la fois une **console interactive** et un **bloc-notes d'essai** : on y évalue des expressions, on y exécute des instructions et on y recueille les traces laissées par le code.

### Évaluer une expression avec « ? »

Le caractère **`?`** est un raccourci signifiant « affiche le résultat ». On tape une expression précédée de `?`, et le résultat s'affiche sur la ligne suivante :

```text
?2 + 2
 4

?Date
 28/05/2026

?DCount("*", "tblClients")
 152

?CurrentDb.Name
 C:\Apps\GestionVentes.accdb
```

C'est le moyen le plus rapide de répondre à des questions comme « que renvoie cette fonction ? », « combien d'enregistrements contient cette table ? » ou « quelle est la valeur de cette propriété à l'instant T ? ». On peut ainsi interroger le **modèle objet** ou les **données** sans écrire la moindre procédure.

### Exécuter des instructions directement

Sans le `?`, ce que l'on tape est **exécuté**. On peut affecter une valeur, appeler une procédure ou une fonction, déclencher une action :

```text
Forms!frmClients!txtRemise = 0.1
MaProcedure
DoCmd.OpenForm "frmClients"
```

Plusieurs instructions peuvent même tenir sur une seule ligne, séparées par le caractère **deux-points** `:`, ce qui autorise de petites boucles d'essai :

```text
For i = 1 To 3 : Debug.Print i : Next
```

La fenêtre devient alors un véritable terrain d'expérimentation : on **teste une ligne** avant de l'intégrer au code, on vérifie un comportement, on déclenche un traitement à la demande.

### Recueillir les sorties de Debug.Print

L'usage symétrique : c'est le **code** qui écrit dans la fenêtre, au moyen de l'instruction **`Debug.Print`**. Tout ce qu'une procédure y envoie s'affiche dans la fenêtre Exécution immédiate — d'où son rôle de **journal de trace** pendant la mise au point.

```vba
Sub Calculer()
    Dim total As Currency
    total = 1500
    Debug.Print "Total calculé : "; total   ' s'affiche dans la fenêtre Exécution immédiate
End Sub
```

La nuance entre les deux est simple : le **`?`** est ce que **vous** tapez interactivement, tandis que **`Debug.Print`** est ce que le **code** utilise ; les deux écrivent dans la même fenêtre. Cet usage de `Debug.Print`, ainsi que l'instruction complémentaire `Debug.Assert`, sont approfondis en section 19.2, dans le cadre du débogage.

### Inspecter et modifier en mode arrêt

La fenêtre Exécution immédiate prend toute sa puissance pendant le **débogage**. Lorsque l'exécution est suspendue (sur un point d'arrêt, par exemple), on peut y **inspecter** la valeur des variables en cours (`?maVariable`), mais aussi les **modifier** à la volée (`maVariable = 10`), appeler des fonctions, et ainsi **tester une correction** sans relancer tout le programme. Cette interaction en direct avec un code figé est l'un des grands atouts du VBE ; le chapitre 19 (points d'arrêt, exécution pas à pas, espions) en détaille la mise en œuvre.

### Quelques points pratiques

- La fenêtre ne conserve qu'un **historique limité** (de l'ordre de quelques centaines de lignes) ; pour la vider, on sélectionne tout et on supprime.
- Elle exécute du code **ligne par ligne** ; les constructions multilignes se ramènent à une seule ligne via le séparateur `:`.
- Les résultats apparaissent **immédiatement**, sans étape de compilation préalable — d'où son nom.

## Deux fenêtres, deux usages

| Fenêtre | Raccourci | Rôle sous Access |
|---|---|---|
| **Propriétés** (éditeur) | <kbd>F4</kbd> | Usage limité : essentiellement renommer les modules |
| **Feuille de propriétés** (Access) | <kbd>F4</kbd> | Toutes les propriétés des formulaires, états et contrôles (chapitres 6, 8) |
| **Exécution immédiate** | <kbd>Ctrl</kbd>+<kbd>G</kbd> | Outil central : évaluer, exécuter, tracer, déboguer |

## À retenir

- La **fenêtre Propriétés de l'éditeur** (<kbd>F4</kbd>) est peu utile sous Access : elle se limite, pour l'essentiel, à la propriété `(Name)` des modules — donc au renommage.
- Elle ne doit **pas être confondue** avec la **feuille de propriétés d'Access** (également <kbd>F4</kbd>, mais côté Access), qui expose, elle, toutes les propriétés des formulaires, états et contrôles (chapitres 6 et 8, annexe B).
- La **fenêtre Exécution immédiate** (<kbd>Ctrl</kbd>+<kbd>G</kbd>) est un outil **central** : le `?` y **affiche** le résultat d'une expression, et toute autre instruction y est **exécutée**.
- Le code y écrit via **`Debug.Print`** ; c'est le journal de trace de la mise au point (approfondi en section 19.2).
- En **mode arrêt**, elle permet d'**inspecter et de modifier** les variables en direct et de tester des corrections (chapitre 19).

---


⏭️ [2.5. Références et bibliothèques COM à activer (DAO, ADO, Scripting...)](/02-interface-environnement/05-references-bibliotheques-com.md)
