🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 16.2. Propriétés (Property Get, Let, Set)

La [section 16.1](01-modules-classe.md) a montré comment créer une classe et l'instancier, en gardant son état dans des champs **privés**. Reste à exposer cet état à l'extérieur — mais de façon **contrôlée**. C'est le rôle des **propriétés** : elles donnent l'apparence d'un simple champ tout en passant par des procédures qui permettent validation, calcul, contrôle des accès en lecture et en écriture. Cette section détaille les trois procédures de propriété — `Get`, `Let` et `Set` — et les usages qu'elles autorisent.

## Pourquoi des propriétés plutôt que des champs publics ?

On pourrait être tenté d'écrire directement `Public Nom As String` dans une classe : de l'extérieur, cela se manipulerait comme une propriété. Mais un champ public est une impasse, car il n'autorise **aucun contrôle**. Une propriété, à l'inverse, permet de :

- **valider** une valeur avant de l'accepter (refuser un âge négatif, un e-mail mal formé) ;
- **contrôler indépendamment** la lecture et l'écriture (rendre une donnée accessible en lecture seule) ;
- **calculer** une valeur à la volée plutôt que de la stocker ;
- offrir une **interface stable** que l'on peut faire évoluer sans tout casser.

C'est l'expression concrète de l'**encapsulation**, principe développé à la [section 16.3](03-methodes-encapsulation.md). La règle est simple : l'état interne reste privé, et on ne l'expose qu'à travers des propriétés.

## Les trois procédures de propriété

Une propriété s'implémente au moyen de trois types de procédures, dont on n'utilise généralement qu'une ou deux à la fois :

- **`Property Get`** — **lit** la valeur. `x = obj.Prop` appelle le `Get`.
- **`Property Let`** — **écrit** une valeur de type simple (nombre, chaîne, date, booléen). `obj.Prop = 42` appelle le `Let`.
- **`Property Set`** — **écrit** une référence d'objet. `Set obj.Prop = unObjet` appelle le `Set`.

La distinction entre `Let` et `Set` reproduit exactement celle, vue en 16.1, entre `=` et `Set =` : `Let` pour les types valeur, `Set` pour les objets. C'est une source d'erreurs fréquente pour qui débute : assigner un objet via une propriété exige `Set`, assigner une valeur simple ne l'exige pas.

## Le patron champ privé + propriété

Le schéma de base associe un **champ privé** (qui stocke réellement la donnée) à un couple `Get`/`Let` (qui en médiatise l'accès).

```vba
' --- Dans clsClient ---
Private m_Nom As String

Public Property Get Nom() As String
    Nom = m_Nom
End Property

Public Property Let Nom(ByVal value As String)
    m_Nom = value
End Property
```

Le champ privé est par convention préfixé `m_` (membre de module). Le nom de la propriété — ici `Nom` — est le nom public. De l'extérieur, l'usage est transparent :

```vba
Dim c As clsClient
Set c = New clsClient
c.Nom = "Dupont"          ' appelle Property Let
Debug.Print c.Nom         ' appelle Property Get
```

## Property Get en détail

`Property Get` se déclare comme une fonction et retourne une valeur en l'affectant au **nom de la propriété** : `Nom = m_Nom`. Son type de retour définit le type de la propriété.

Détail important lorsque la propriété renvoie un **objet** : on retourne alors la référence avec `Set`, exactement comme partout ailleurs :

```vba
Public Property Get Adresse() As clsAdresse
    Set Adresse = m_Adresse        ' Set, car on retourne un objet
End Property
```

## Property Let en détail

`Property Let` reçoit la valeur affectée comme **dernier paramètre**. On la déclare en `ByVal` — on ne veut pas modifier la variable de l'appelant. Le type de ce paramètre doit correspondre au type de retour du `Get`.

```vba
Public Property Let Nom(ByVal value As String)
    m_Nom = value
End Property
```

`Let` ne s'emploie que pour les **types valeur**. Pour un objet, c'est `Set` qu'il faut.

## Property Set en détail

`Property Set` est l'équivalent de `Let` pour les **références d'objet**. On l'appelle via `Set obj.Prop = unObjet`. Une propriété qui expose un objet se dote donc d'un couple `Get` + `Set` :

```vba
Private m_Adresse As clsAdresse

Public Property Get Adresse() As clsAdresse
    Set Adresse = m_Adresse
End Property

Public Property Set Adresse(ByVal value As clsAdresse)
    Set m_Adresse = value
End Property
```

Côté utilisation :

```vba
Set c.Adresse = New clsAdresse     ' appelle Property Set
```

## Propriétés en lecture seule ou en écriture seule

C'est l'un des grands avantages des propriétés : on contrôle séparément la lecture et l'écriture, simplement en fournissant ou non chaque procédure.

Une propriété **en lecture seule** ne définit qu'un `Get`, sans `Let` ni `Set`. Toute tentative d'affectation provoque alors une erreur de compilation. C'est typiquement le cas d'un identifiant qui ne doit pas changer après création :

```vba
Private m_Id As Long

Public Property Get Id() As Long      ' lecture seule : aucun Let
    Id = m_Id
End Property
```

Le champ `m_Id` est alors renseigné **depuis l'intérieur** de la classe — dans `Class_Initialize`, par une méthode `Init`, ou par une fabrique ([section 16.7](07-singleton-factory-observer.md)) — et reste en lecture seule pour le monde extérieur.

À l'inverse, une propriété **en écriture seule** ne définit qu'un `Let` (ou `Set`), sans `Get`. C'est plus rare, mais utile par exemple pour un mot de passe que l'on accepte d'écrire mais jamais de relire.

## Propriétés calculées

Une propriété n'est pas obligée de refléter un champ stocké : son `Get` peut **calculer** sa valeur à la volée. Une telle propriété n'a ni champ privé ni `Let`.

```vba
Public Property Get NomComplet() As String
    NomComplet = Trim(m_Prenom & " " & m_Nom)
End Property
```

`NomComplet` est dérivée de `Prenom` et `Nom` ; elle est toujours cohérente avec eux, et n'occupe aucun stockage. C'est un usage puissant : exposer de l'information dérivée comme s'il s'agissait d'une donnée, tout en la maintenant automatiquement à jour.

## Valider dans Property Let

La validation est la justification la plus forte des propriétés. Le `Let` (ou `Set`) peut **contrôler la valeur avant de l'accepter** et lever une erreur en cas de problème, au moyen de `Err.Raise` (cf. [section 13.8](/13-gestion-erreurs/08-erreurs-personnalisees-raise.md)).

```vba
Public Property Let Email(ByVal value As String)
    If InStr(value, "@") = 0 Then
        Err.Raise vbObjectError + 513, "clsClient.Email", _
                  "Adresse e-mail invalide : " & value
    End If
    m_Email = value
End Property
```

Ici, un objet `clsClient` ne pourra **jamais** contenir un e-mail dépourvu d'arobase : la classe garantit elle-même la cohérence de son état, où que vienne l'affectation. C'est exactement ce qu'un champ public serait incapable d'assurer.

## Propriétés paramétrées

Une propriété peut aussi accepter des **arguments**, devenant une propriété paramétrée. Cela permet notamment d'indexer un accès, comme la propriété `Item` d'une collection : `Property Get Item(ByVal index As Long) As clsX`. Ce mécanisme est au cœur des collections personnalisées, traitées à la [section 16.5](05-collections-interfaces.md).

## Points de vigilance

- **`Let` pour les valeurs, `Set` pour les objets** : assigner un objet via une propriété exige `Set`.
- **Dans un `Get` qui renvoie un objet, utiliser `Set`** (`Set Prop = m_objet`), pas une affectation simple.
- **Types cohérents** : le type de retour du `Get` et le type du paramètre du `Let`/`Set` doivent correspondre.
- **`ByVal` pour le paramètre des setters** : ne pas modifier la variable de l'appelant.
- **Lecture seule = `Get` sans `Let`/`Set`** : le moyen propre de protéger une donnée (identité, valeur calculée).
- **Valider dans le setter** : c'est l'avantage décisif des propriétés sur les champs publics ; lever une erreur (`Err.Raise`) sur une valeur invalide.
- **Préférer les propriétés aux champs publics** : même pour une simple lecture/écriture, elles préservent l'évolutivité et l'encapsulation.

## En résumé

Les propriétés exposent l'état d'un objet de manière **contrôlée**, à travers trois procédures : `Property Get` lit la valeur (et renvoie un objet avec `Set` si nécessaire), `Property Let` écrit une valeur simple, et `Property Set` écrit une référence d'objet — la distinction `Let`/`Set` reflétant celle entre `=` et `Set =`. Le patron usuel associe un champ privé `m_` à un couple `Get`/`Let`. Selon les procédures fournies, on obtient des propriétés en **lecture seule** (`Get` seul), en **écriture seule**, ou **calculées** (un `Get` qui dérive sa valeur sans champ stocké). Surtout, le setter peut **valider** la valeur et refuser une donnée incohérente via `Err.Raise` — ce qu'un champ public ne permet jamais, et qui fait des propriétés le bon choix par défaut pour exposer l'état d'une classe.

La section suivante complète les objets par leurs comportements et formalise le principe sous-jacent : les [méthodes et l'encapsulation](03-methodes-encapsulation.md).

⏭️ [16.3. Méthodes et encapsulation](/16-poo-vba-access/03-methodes-encapsulation.md)
