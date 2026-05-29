🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 1.4. Activation et paramétrage des options de développement

La section précédente s'est conclue sur la question de la sécurité, qui conditionne l'exécution du code. Avant d'aborder le développement proprement dit, il faut justement **préparer son environnement** : autoriser l'exécution du code, activer les aides à l'écriture et régler quelques options qui feront gagner un temps précieux tout en évitant des erreurs. C'est l'objet de cette section.

Un point mérite d'être noté d'emblée, dans le prolongement de la section 1.2 : contrairement à Excel et Word, **Access ne possède pas d'onglet « Développeur » à activer**. Les outils de développement y sont déjà présents — sur le ruban **Créer** (groupe **Macros et code**) et sur l'onglet **Outils de base de données**. « Activer les options de développement » consiste donc surtout à accéder à l'éditeur de code, à autoriser son exécution et à le configurer correctement.

## Accéder à l'environnement de code

Pour ouvrir l'éditeur VBA (le VBE, *Visual Basic Editor*), trois voies équivalentes :

- le raccourci <kbd>Alt</kbd>+<kbd>F11</kbd> (le plus rapide) ;
- le bouton **Visual Basic** du ruban **Créer**, groupe **Macros et code** ;
- le bouton **Visual Basic** de l'onglet **Outils de base de données**.

Le même ruban **Créer** permet aussi de créer directement un **Module** (module standard) ou un **Module de classe**. L'accès à l'éditeur et la prise en main de son interface sont détaillés au chapitre 2 (section 2.1) ; nous nous concentrons ici sur sa **configuration**.

## Les options de l'éditeur VBA

Une fois dans le VBE, ouvrez **Outils > Options**. C'est ici que se trouvent les réglages les plus importants pour écrire du code confortablement et sûrement.

### Déclaration des variables obligatoire (le réglage essentiel)

Dans l'onglet **Éditeur**, cochez **Déclaration des variables obligatoire** (*Require Variable Declaration*). C'est, de loin, le paramètre le plus important. Une fois activé, Access insère automatiquement la ligne `Option Explicit` en tête de **chaque nouveau module**.

```vba
Option Compare Database
Option Explicit   ' ajouté automatiquement quand « Déclaration des variables obligatoire » est coché
```

`Option Explicit` impose de **déclarer toutes les variables** avant de les utiliser. Sans cette contrainte, une simple faute de frappe dans un nom de variable crée silencieusement une nouvelle variable vide, source de bugs particulièrement difficiles à traquer :

```vba
Dim montantTotal As Currency
montantTotal = 100
montanTotal = montantTotal + 50   ' faute de frappe : sans Option Explicit, créée en douce et = 50
' montantTotal vaut toujours 100, mais aucune erreur n'est signalée
```

Avec `Option Explicit`, ce code provoque une erreur de compilation immédiate (« Variable non définie »), pointant exactement la ligne fautive. **Attention** : cocher cette option n'agit que sur les modules *créés ensuite* ; pour les modules existants, ajoutez `Option Explicit` à la main en tête de chacun.

### Les aides à la saisie

Toujours dans l'onglet **Éditeur**, plusieurs options assistent l'écriture et méritent d'être activées :

- **Vérification automatique de la syntaxe** (*Auto Syntax Check*) : signale immédiatement une ligne mal formée. (Certains la désactivent pour éviter les boîtes de dialogue intempestives ; la coloration en rouge de la ligne suffit alors à repérer le problème — affaire de préférence.)
- **Liste des composants automatique** (*Auto List Members*) : affiche la liste des propriétés et méthodes d'un objet dès que l'on tape le point. Indispensable pour explorer le modèle objet sans tout mémoriser.
- **Info express automatique** (*Auto Quick Info*) : montre la signature d'une procédure ou d'une fonction (ses arguments) pendant la frappe.
- **Info-bulles automatiques** (*Auto Data Tips*) : en mode arrêt (débogage), affiche la valeur d'une variable au survol de la souris.
- **Indentation automatique** (*Auto Indent*) et **largeur de tabulation** : conservent une présentation lisible du code.

### Mise en forme du code

L'onglet **Format de l'éditeur** permet de choisir police et couleurs de la coloration syntaxique. Une police à chasse fixe (*Consolas*, *Courier New*) améliore nettement la lisibilité ; le reste relève du confort personnel.

### Interception des erreurs et compilation

L'onglet **Général** contient un réglage qui surprend souvent les débutants : la section **Interception des erreurs** (*Error Trapping*), avec trois modes.

- **Arrêt sur les erreurs non gérées** (*Break on Unhandled Errors*) : mode par défaut recommandé. Le code ne s'arrête que sur les erreurs qu'aucun gestionnaire `On Error` ne prend en charge.
- **Arrêt sur toutes les erreurs** (*Break on All Errors*) : interrompt à *chaque* erreur, même celles qui sont gérées. Très utile **temporairement** lorsqu'un bug est « avalé » par un `On Error Resume Next` et que l'on cherche à le localiser.
- **Arrêt dans les modules de classe** (*Break in Class Module*) : arrête précisément dans le module de classe à l'origine d'une erreur non gérée, plutôt qu'au point d'appel — pratique lors du développement de classes.

En pratique : laissez **Arrêt sur les erreurs non gérées** au quotidien, et basculez ponctuellement sur **Arrêt sur toutes les erreurs** pour débusquer une erreur masquée. Ces mécanismes sont approfondis aux chapitres 13 (gestion des erreurs) et 19 (débogage).

Le même onglet propose **Compiler à la demande** et **Compilation en arrière-plan** : les réglages par défaut conviennent dans la grande majorité des cas. Pensez toutefois à compiler régulièrement votre projet (menu **Débogage > Compiler**) pour détecter au plus tôt les erreurs de compilation.

## Permettre l'exécution du code : la sécurité

Comme l'a montré la section 1.3, **le code VBA ne s'exécute pas tant que la base n'est pas approuvée**. Si votre base s'ouvre avec une barre de message d'avertissement et que rien ne se déclenche, c'est presque toujours la cause. Le réglage se fait via **Fichier > Options > Centre de gestion de la confidentialité > Paramètres du Centre de gestion de la confidentialité**.

Deux endroits y sont pertinents :

- **Paramètres des macros** : déterminent le comportement par défaut à l'ouverture d'une base non approuvée. Évitez l'option « Activer toutes les macros » de façon permanente : elle est commode mais dangereuse.
- **Emplacements approuvés** (*Trusted Locations*) : c'est la bonne pratique pour le développement. Ajoutez-y le **dossier de travail** contenant vos bases : tout fichier qui s'y trouve sera approuvé automatiquement, et le code s'exécutera sans avertissement à chaque ouverture.

Le modèle de sécurité d'Access et ces paramètres sont traités en détail en section 20.5.

## Les options d'Access utiles au développement

En dehors du VBE, plusieurs réglages se trouvent dans **Fichier > Options**. Plusieurs catégories concernent directement le développeur.

### Base de données active

La catégorie **Base de données active** (*Current Database*) regroupe des options propres au fichier ouvert :

- **Utiliser les touches spéciales Access** (*Use Access Special Keys*) : laissez cette option **activée pendant le développement** ; elle autorise notamment <kbd>Alt</kbd>+<kbd>F11</kbd> et l'affichage du volet de navigation. On la désactive souvent au moment de *verrouiller* l'interface d'une application livrée (voir section 17.7), mais pas avant.
- **Compacter lors de la fermeture** (*Compact on Close*) : compacte la base à chaque fermeture. Pratique en développement, où le fichier gonfle vite ; à manier avec prudence sur un fichier partagé (voir section 18.7 sur le compactage).
- **Options de correction automatique de nom** (*Name AutoCorrect Options*) : voir l'encadré ci-dessous.

### Le cas de la correction automatique de nom

La fonction **Correction automatique de nom** tente de répercuter automatiquement les renommages d'objets (par exemple un champ renommé) dans les objets qui en dépendent. En théorie commode, elle est très largement **déconseillée** par les développeurs expérimentés : elle alourdit la base, peut nuire aux performances et masque parfois de vraies dépendances cassées. La recommandation courante est de **décocher** « Suivre les informations de correction automatique de nom » et « Effectuer la correction automatique de nom », puis de gérer les renommages de façon maîtrisée. C'est un choix discuté, mais qui fait consensus dans le développement sérieux.

### Concepteurs d'objets

La catégorie **Concepteurs d'objets** (*Object Designers*) contient les options de **vérification des erreurs** dans les formulaires et états (les petits indicateurs signalant, par exemple, des contrôles non liés ou des étiquettes orphelines). Les laisser activées aide à repérer les anomalies dès la conception.

### Paramètres du client

La catégorie **Paramètres du client** (*Client Settings*) expose des réglages avancés, dont certains touchent au multi-utilisateur :

- **Mode d'ouverture par défaut** (*Default open mode*) : partagé ou exclusif ;
- **Verrouillage des enregistrements par défaut** (*Default record locking*).

Ces paramètres prennent toute leur importance en environnement partagé ; ils sont traités au chapitre 15 (multi-utilisateurs et verrouillage).

## Afficher les objets cachés et système

Pendant le développement, il est parfois utile de voir des objets normalement masqués. Un clic droit sur la barre de titre du **volet de navigation > Options de navigation** permet d'afficher les **objets cachés** et les **objets système**. À réserver au développement : ces objets sont masqués pour de bonnes raisons en exploitation.

## 32 bits ou 64 bits : un point de vigilance

Access existe en versions **32 bits** et **64 bits**, et cette différence a une incidence dès que l'on **déclare des API Windows** : la syntaxe des déclarations (`Declare`, mot-clé `PtrSafe`, type `LongPtr`) en dépend. Il est donc utile de savoir, dès maintenant, sous quelle architecture vous travaillez. La question est traitée en détail en section 21.7 ; retenez simplement, à ce stade, qu'elle existe et qu'elle peut affecter la portabilité du code entre postes.

> ℹ️ Les **références et bibliothèques COM** à activer (DAO, ADO, Scripting…) constituent un autre volet de la configuration de l'environnement. Elles sont traitées spécifiquement en section 2.5.

## Configuration recommandée — récapitulatif

| Paramètre | Emplacement | Réglage recommandé |
|---|---|---|
| Déclaration des variables obligatoire | VBE > Outils > Options > Éditeur | **Activé** |
| Liste des composants / Info express auto | VBE > Outils > Options > Éditeur | Activées |
| Info-bulles automatiques (débogage) | VBE > Outils > Options > Éditeur | Activée |
| Interception des erreurs | VBE > Outils > Options > Général | Arrêt sur les erreurs non gérées |
| Emplacement approuvé (dossier de dev) | Fichier > Options > Centre de gestion… > Emplacements approuvés | **Ajouté** |
| Touches spéciales Access | Fichier > Options > Base de données active | Activées (pendant le dev) |
| Compacter lors de la fermeture | Fichier > Options > Base de données active | Selon contexte (pratique en dev) |
| Correction automatique de nom | Fichier > Options > Base de données active | **Désactivée** (recommandé) |
| Vérification des erreurs (conception) | Fichier > Options > Concepteurs d'objets | Activée |

## À retenir

- Access **n'a pas d'onglet « Développeur »** à activer : les outils sont déjà sur **Créer** et **Outils de base de données**, et l'éditeur s'ouvre par <kbd>Alt</kbd>+<kbd>F11</kbd>.
- Le réglage le plus important est **Déclaration des variables obligatoire**, qui ajoute `Option Explicit` et évite quantité de bugs liés aux fautes de frappe.
- Activez les **aides à la saisie** (liste des composants, info express, info-bulles) pour écrire plus vite et plus sûrement.
- Le **code ne s'exécute que dans une base approuvée** : déclarer son dossier de travail comme **emplacement approuvé** est la bonne pratique de développement.
- Quelques options d'Access méritent attention : **touches spéciales** (activées en dev), **correction automatique de nom** (à désactiver), **vérification des erreurs**, et les paramètres de **mode d'ouverture / verrouillage** pour le multi-utilisateur.
- Sachez sous quelle architecture (**32 ou 64 bits**) vous travaillez : cela influe sur les déclarations d'API Windows.

---


⏭️ [1.5. Structure d'une application Access (tables, requêtes, formulaires, états, modules)](/01-introduction-vba-access/05-structure-application-access.md)
