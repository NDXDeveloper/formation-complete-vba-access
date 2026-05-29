🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 16.5. Collections personnalisées et interfaces (Implements)

Cette section réunit deux outils de structuration qui se complètent. Les **collections personnalisées** permettent de regrouper des objets de façon typée et signifiante, là où la `Collection` native de VBA se contente de tout accepter sans contrôle. Les **interfaces**, via le mot-clé `Implements`, apportent le **polymorphisme** — la capacité de traiter uniformément des objets de classes différentes — qui, en l'absence d'héritage en VBA, est le moyen privilégié d'abstraire et de découpler. Les deux se rejoignent naturellement, comme on le verra en fin de section.

## Les collections personnalisées

### Pourquoi une collection personnalisée ?

La `Collection` intégrée de VBA stocke n'importe quoi : ses éléments sont des `Variant`, sans aucun contrôle de type. On peut y ajouter par erreur un objet du mauvais type, et à la lecture, on récupère un `Variant` sans assistance à la saisie.

Une **collection personnalisée** est une classe qui **enveloppe** une `Collection` privée et expose une interface **typée** : un `Add` qui n'accepte qu'un type précis, un `Item` qui retourne ce type, et un point naturel où ajouter de la logique métier (interdire les doublons, trier, filtrer). On gagne ainsi la sûreté de typage, l'assistance à la saisie, et la maîtrise du comportement de la collection.

### Le patron : envelopper une Collection

Le principe consiste à garder la vraie collection en champ privé et à n'exposer que des membres typés qui lui délèguent le travail.

```vba
' --- Module de classe : clsClients (collection personnalisée) ---
Option Explicit

Private m_Items As Collection           ' stockage réel, privé

Private Sub Class_Initialize()
    Set m_Items = New Collection
End Sub

Public Sub Add(ByVal client As clsClient)
    m_Items.Add client, CStr(client.Id)  ' clé = Id, pour un accès direct
End Sub

Public Property Get Item(ByVal index As Variant) As clsClient
    Set Item = m_Items(index)            ' par position (1..Count) ou par clé
End Property

Public Property Get Count() As Long
    Count = m_Items.Count
End Property

Public Sub Remove(ByVal index As Variant)
    m_Items.Remove index
End Sub
```

`Add` n'accepte qu'un `clsClient` ; `Item` (une propriété paramétrée, cf. [section 16.2](02-proprietes-get-let-set.md)) en retourne un. La collection est typée de bout en bout, et l'on pourrait enrichir `Add` d'un contrôle de doublon ou d'une validation.

### Rendre la collection énumérable (For Each)

En l'état, on parcourt cette collection par index :

```vba
Dim i As Long
For i = 1 To mesClients.Count
    Debug.Print mesClients.Item(i).Nom
Next i
```

Pour autoriser la syntaxe `For Each`, il faut exposer un membre spécial d'énumération. La technique consiste à ajouter une fonction `NewEnum` qui délègue à l'énumérateur caché de la collection interne, **et à lui affecter l'identifiant de membre `-4`** (la valeur réservée à l'énumération) :

```vba
Public Function NewEnum() As IUnknown
    Set NewEnum = m_Items.[_NewEnum]
End Function
' Attribut requis (non visible dans l'éditeur) : Attribute NewEnum.VB_UserMemId = -4
```

Cet attribut ne se définit pas directement dans l'éditeur VBA : il faut **exporter le module de classe**, ajouter la ligne `Attribute NewEnum.VB_UserMemId = -4` dans le fichier `.cls`, puis le réimporter — ou utiliser un complément comme Rubberduck qui automatise l'opération. De la même manière, affecter l'identifiant `0` à la propriété `Item` (`Attribute Item.VB_UserMemId = 0`) en fait le **membre par défaut**, autorisant l'écriture abrégée `mesClients(1)` pour `mesClients.Item(1)`. Ces réglages sont facultatifs : une collection parcourue par index reste parfaitement utilisable sans eux.

### Collection ou Dictionary ?

Une collection personnalisée peut s'appuyer sur une `Collection` (ordonnée, indexée à partir de 1, clé optionnelle) ou sur un `Scripting.Dictionary` (couples clé/valeur, méthode `Exists`, gestion plus souple des clés). Le choix dépend des besoins — accès par position, test d'existence rapide, mise à jour par clé — selon les caractéristiques rappelées à la [section 3.7](/03-rappels-fondamentaux/07-collections-dictionary.md).

## Les interfaces (Implements)

### Qu'est-ce qu'une interface ?

Une **interface** est un **contrat** : un ensemble de propriétés et de méthodes — leurs signatures seulement, sans implémentation — qu'une classe s'engage à fournir. En VBA, on définit une interface comme un module de classe ne contenant que les **signatures** des membres, avec des corps vides. La convention préfixe son nom par `I`.

```vba
' --- Module de classe : IExportable (l'interface) ---
Option Explicit

Public Sub Exporter(ByVal chemin As String)
    ' aucun corps : un contrat, pas une implémentation
End Sub

Public Property Get Libelle() As String
End Property
```

### Pourquoi des interfaces ? Le polymorphisme sans héritage

Comme rappelé en introduction du chapitre, VBA **ne propose pas d'héritage d'implémentation**. Les interfaces comblent ce manque sur le plan du **polymorphisme** : elles permettent à du code de manipuler n'importe quel objet respectant un contrat donné, sans connaître sa classe concrète. On ne partage pas de code par héritage, mais on partage un **contrat** par interface — et l'on traite uniformément tous ceux qui l'honorent.

C'est aussi le fondement de l'**abstraction** : du code métier peut dépendre d'une interface plutôt que d'une implémentation précise, ce qui est exactement ce dont a besoin le patron Repository pour rendre l'application indépendante du moteur de données ([section 16.6](06-patron-repository.md)).

### Implémenter une interface

Une classe déclare qu'elle honore un contrat avec `Implements`, puis fournit chacun des membres de l'interface sous la forme `NomInterface_Membre`.

```vba
' --- En tête de clsClient ---
Implements IExportable

' Implémentation des membres de l'interface (préfixe IExportable_)
Private Sub IExportable_Exporter(ByVal chemin As String)
    ' export réel d'un client vers 'chemin'
End Sub

Private Property Get IExportable_Libelle() As String
    IExportable_Libelle = "Client : " & m_Nom
End Property
```

Trois points à retenir : la classe **doit** implémenter **tous** les membres de l'interface, sous peine de ne pas compiler ; ces implémentations sont typiquement `Private`, car on y accède via l'interface et non directement ; et une classe peut implémenter plusieurs interfaces (plusieurs `Implements`) en plus de ses propres membres publics. Les listes déroulantes de l'éditeur VBA aident à générer les ossatures : il suffit de sélectionner l'interface dans la liste de gauche.

### Utiliser le polymorphisme

L'intérêt apparaît à l'usage : on déclare une variable du **type de l'interface**, on lui affecte n'importe quel objet qui l'implémente, et l'on appelle les membres du contrat — la classe concrète fournissant l'exécution.

```vba
Dim exp As IExportable
Set exp = New clsClient            ' un clsClient EST AUSSI un IExportable
exp.Exporter "C:\clients.txt"      ' exécute IExportable_Exporter de clsClient

' Traiter uniformément des objets de classes différentes :
Dim o As IExportable
For Each o In lesExportables        ' une collection d'objets IExportable
    Debug.Print o.Libelle
    o.Exporter cheminBase & o.Libelle & ".txt"
Next o
```

Si `clsClient` et `clsCommande` implémentent tous deux `IExportable`, le même code les traite indifféremment. On peut aussi tester l'appartenance à un contrat avec `TypeOf` :

```vba
If TypeOf unObjet Is IExportable Then
    Dim e As IExportable
    Set e = unObjet
    e.Exporter "..."
End If
```

Le code dépend ici de l'**abstraction** (`IExportable`), non des classes concrètes — ce qui le rend extensible : ajouter une nouvelle classe exportable ne demande aucune modification du code qui consomme l'interface.

## Synthèse : une collection d'objets polymorphes

Les deux notions se combinent naturellement. Une collection personnalisée peut être typée non pas sur une classe concrète, mais sur une **interface** : une `clsExportables` dont l'`Item` retourne un `IExportable`. On obtient alors une collection homogène en apparence, mais contenant des objets de classes variées, tous traités à travers leur contrat commun. C'est un assemblage courant dans une conception orientée objet soignée — et un avant-goût des patrons de la section suivante.

## Points de vigilance

- **Collection personnalisée = enveloppe typée** : garder la `Collection` interne privée, n'exposer que des membres typés.
- **`For Each` exige le membre `NewEnum`** avec l'identifiant `-4`, défini en éditant le `.cls` exporté (ou via Rubberduck) ; à défaut, parcourir par index.
- **Une interface ne contient que des signatures** : corps vides, préfixe `I` par convention.
- **`Implements` oblige à fournir tous les membres** de l'interface, nommés `NomInterface_Membre`, généralement `Private`.
- **Le polymorphisme passe par une variable typée sur l'interface** : déclarer `As IInterface`, affecter n'importe quel objet conforme.
- **Tester un contrat avec `TypeOf … Is IInterface`** avant de l'utiliser, en cas de doute sur le type.
- **Dépendre des interfaces, pas des classes concrètes** : c'est ce qui rend le code extensible et découplé (clé du Repository).

## En résumé

Une **collection personnalisée** est une classe qui enveloppe une `Collection` (ou un `Dictionary`) privée et expose une interface **typée** — `Add`, `Item`, `Count`, `Remove` — apportant sûreté de typage et logique métier ; pour autoriser `For Each`, on y ajoute un membre `NewEnum` portant l'identifiant `-4`, défini en éditant le fichier `.cls`. Une **interface** est un contrat de signatures, défini comme un module de classe aux corps vides ; une classe l'honore avec `Implements` en fournissant tous ses membres sous la forme `NomInterface_Membre`. Les interfaces apportent le **polymorphisme** que l'absence d'héritage interdit autrement : du code déclaré sur le type d'une interface manipule uniformément tout objet qui l'implémente, sans connaître sa classe concrète. Cette dépendance à l'abstraction plutôt qu'aux implémentations est le ressort du découplage — et, combinée aux collections typées, le socle des patrons de conception qui suivent.

La section suivante applique ces idées à un patron majeur : le [patron Repository, pour abstraire l'accès aux données](06-patron-repository.md).

⏭️ [16.6. Patron Repository — abstraction de l'accès aux données](/16-poo-vba-access/06-patron-repository.md)
