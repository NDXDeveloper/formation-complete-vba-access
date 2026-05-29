🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 16.4. Événements personnalisés dans une classe (RaiseEvent / WithEvents)

Au-delà des événements intégrés `Class_Initialize` et `Class_Terminate` ([section 16.1](01-modules-classe.md)), une classe VBA peut **définir ses propres événements**. Un objet devient alors capable de **notifier** d'autres parties du programme lorsqu'il se passe quelque chose — sans rien savoir de qui l'écoute. C'est le mécanisme de publication/abonnement de VBA, déjà familier au travers des formulaires et contrôles, et que cette section apprend à exploiter dans vos propres classes.

## L'idée : publication et abonnement

Un objet qui déclenche des événements joue le rôle de **publieur** : il annonce qu'un fait s'est produit. D'autres morceaux de code, les **abonnés** (ou écouteurs), s'inscrivent pour être prévenus et réagir. Le publieur, lui, ignore tout des abonnés : il se contente d'émettre.

C'est précisément ce que vous connaissez déjà des formulaires : un bouton déclenche `Click`, un formulaire `AfterUpdate`, et votre code y réagit (chapitre 8). La nouveauté est de pouvoir faire la même chose avec vos classes métier — un compte qui signale un découvert, un traitement qui annonce sa progression.

## Les trois pièces

Mettre en place un événement personnalisé suppose trois éléments :

1. **déclarer** l'événement dans la classe publieuse, avec `Event` ;
2. **déclencher** l'événement dans cette même classe, avec `RaiseEvent` ;
3. **s'abonner** depuis l'écouteur, en déclarant l'objet `WithEvents` et en écrivant des gestionnaires.

## Déclarer un événement (côté publieur)

On déclare un événement au niveau du module de classe, avec le mot-clé `Event`, en lui donnant une signature :

```vba
' --- Module de classe : clsCompte (le publieur) ---
Option Explicit

Private m_Solde As Currency

' Déclaration des événements
Public Event SoldeModifie(ByVal ancien As Currency, ByVal nouveau As Currency)
Public Event Decouvert(ByVal montant As Currency)
```

Un événement peut transporter des **paramètres** vers ses abonnés, en `ByVal` ou en `ByRef`. Il ne **retourne jamais de valeur** : ce n'est pas une fonction, mais une notification.

## Déclencher un événement (côté publieur)

On émet l'événement avec `RaiseEvent`, depuis l'intérieur de la classe qui l'a déclaré. Si aucun abonné n'écoute, l'appel est sans effet — c'est parfaitement sûr.

```vba
Public Property Get Solde() As Currency
    Solde = m_Solde
End Property

Public Sub Crediter(ByVal montant As Currency)
    ModifierSolde m_Solde + montant
End Sub

Public Sub Debiter(ByVal montant As Currency)
    ModifierSolde m_Solde - montant
End Sub

Private Sub ModifierSolde(ByVal nouveau As Currency)
    Dim ancien As Currency
    ancien = m_Solde
    m_Solde = nouveau
    RaiseEvent SoldeModifie(ancien, m_Solde)            ' notifie le changement
    If m_Solde < 0 Then RaiseEvent Decouvert(m_Solde)   ' notifie le découvert
End Sub
```

## S'abonner (côté écouteur)

Pour recevoir les événements d'un objet, l'écouteur déclare la variable qui le référence avec le mot-clé **`WithEvents`**. Cette déclaration se fait **au niveau module**, dans un module de classe ou un module de formulaire/état.

```vba
' --- Dans un module de formulaire ou une classe (l'écouteur) ---
Option Explicit

Private WithEvents m_Compte As clsCompte    ' déclaration au niveau module

Private Sub Form_Load()
    Set m_Compte = New clsCompte            ' à partir d'ici, on écoute
    m_Compte.Crediter 100
    m_Compte.Debiter 250                    ' déclenchera SoldeModifie puis Decouvert
End Sub

' Gestionnaires : nom = nomVariable_NomEvenement, même signature que l'événement
Private Sub m_Compte_SoldeModifie(ByVal ancien As Currency, ByVal nouveau As Currency)
    Debug.Print "Solde : " & ancien & " -> " & nouveau
End Sub

Private Sub m_Compte_Decouvert(ByVal montant As Currency)
    MsgBox "Découvert : " & Format(montant, "Currency"), vbExclamation
End Sub
```

Une fois la variable déclarée `WithEvents`, elle apparaît dans les listes déroulantes de l'éditeur VBA, qui génèrent les ossatures des gestionnaires. Le nom de chaque gestionnaire suit le format `nomVariable_NomEvenement`, et sa signature doit correspondre exactement à celle de l'événement. C'est l'affectation `Set m_Compte = New clsCompte` qui « branche » l'écoute.

## Les contraintes de WithEvents

`WithEvents` impose plusieurs règles qu'il faut connaître :

- la déclaration doit être faite **au niveau module**, jamais à l'intérieur d'une procédure ;
- elle n'est possible que dans un **module de classe** ou de formulaire/état — **pas dans un module standard**, car les gestionnaires d'événements ont besoin d'une instance qui les héberge ;
- la variable doit être typée avec une **classe précise** (liaison anticipée) : on ne peut pas utiliser `WithEvents` avec un type `Object` en liaison tardive ;
- une variable `WithEvents` ne référence **qu'un objet à la fois** ; on peut la repointer vers une autre instance avec `Set`, et les gestionnaires réagiront alors à ce nouvel objet.

Pour écouter plusieurs objets de la même classe simultanément, il faut une collection d'écouteurs — un cas qui rejoint le patron Observer ([section 16.7](07-singleton-factory-observer.md)).

## Paramètres ByRef : le motif d'annulation

Un événement peut comporter un paramètre **`ByRef`** que l'abonné modifie, renvoyant ainsi une information au publieur. Le cas emblématique est un paramètre d'annulation — exactement comme le `Cancel` des événements de formulaire. Le publieur émet l'événement, l'abonné peut poser un veto, et le publieur en tient compte.

```vba
' Côté publieur :
Public Event AvantDebit(ByVal montant As Currency, ByRef annuler As Boolean)

Public Sub Debiter(ByVal montant As Currency)
    Dim annuler As Boolean
    RaiseEvent AvantDebit(montant, annuler)     ' un abonné peut renseigner 'annuler'
    If annuler Then Exit Sub                     ' le publieur respecte le veto
    ModifierSolde m_Solde - montant
End Sub
```

```vba
' Côté écouteur :
Private Sub m_Compte_AvantDebit(ByVal montant As Currency, ByRef annuler As Boolean)
    If montant > 1000 Then annuler = True        ' veto sur les gros débits
End Sub
```

Ce motif permet à un abonné d'**influencer** le déroulement d'une action, et non seulement d'en être informé.

## Pourquoi des événements ? Le découplage

L'intérêt majeur des événements est le **découplage**. Le publieur n'a aucune dépendance envers ses abonnés : il annonce « ceci s'est produit » sans connaître ni l'identité ni le nombre des écouteurs. On peut ajouter ou retirer des abonnés **sans toucher au publieur**. Un même publieur peut d'ailleurs avoir **plusieurs abonnés**, chacun détenant sa propre référence `WithEvents` vers lui : tous reçoivent l'événement.

Ce couplage faible est le fondement du patron **Observer** (section 16.7), et un excellent moyen de laisser des objets métier notifier l'interface — par exemple un traitement long qui émet des événements de progression qu'un formulaire affiche dans une barre de progression ([section 17.5](/17-interface-utilisateur-avancee/05-barre-progression.md)).

## Événements ou appel direct ?

L'alternative à un événement est l'appel direct : le publieur appellerait lui-même une méthode de l'écouteur. Mais cela suppose que le publieur **connaisse** l'écouteur — un couplage fort, qui lie les deux et empêche d'en changer librement. L'événement, à l'inverse, laisse le publieur indépendant. On préfère donc l'événement lorsque le publieur ne doit pas dépendre des écouteurs, ou lorsque ceux-ci peuvent être absents ou multiples.

## Cycle de vie et pièges

La variable `WithEvents` doit **vivre** aussi longtemps que l'on souhaite écouter : c'est pourquoi elle est un champ de niveau module de l'écouteur. Si elle sort de portée, l'écoute cesse.

Côté références, c'est l'**écouteur qui référence le publieur** (via sa variable `WithEvents`), et non l'inverse : tant qu'un écouteur tient cette référence, le publieur n'est pas détruit. Pour cesser d'écouter, on remet la variable à `Nothing` (`Set m_Compte = Nothing`). Comme le publieur ne référence pas ses écouteurs, ce mécanisme ne crée pas de référence circulaire.

## Points de vigilance

- **Trois pièces** : `Event` (déclarer), `RaiseEvent` (déclencher, depuis la classe), `WithEvents` (s'abonner).
- **`WithEvents` au niveau module, dans une classe ou un formulaire** — jamais dans un module standard ni dans une procédure.
- **Liaison anticipée obligatoire** : `WithEvents` exige un type de classe précis, pas `Object`.
- **Gestionnaire = `nomVariable_NomEvenement`** avec une signature identique à l'événement.
- **`Set` pour brancher l'écoute, `Nothing` pour l'arrêter** : la variable doit vivre tant qu'on écoute.
- **Une variable `WithEvents` = un objet à la fois** ; plusieurs objets → collection d'écouteurs (Observer, 16.7).
- **Paramètre `ByRef`** pour qu'un abonné renvoie une information (motif d'annulation).

## En résumé

Une classe VBA peut définir ses propres **événements** pour notifier d'autres objets, selon un schéma de publication/abonnement. Trois éléments l'organisent : la classe publieuse **déclare** l'événement (`Public Event Nom(params)`) et le **déclenche** (`RaiseEvent Nom(args)`), tandis que l'écouteur déclare l'objet `WithEvents` au niveau module et écrit des gestionnaires nommés `variable_Evenement`, branchés par un `Set`. `WithEvents` impose ses règles — niveau module, hôte de type classe ou formulaire, liaison anticipée, un objet à la fois. Un paramètre `ByRef` permet à un abonné d'influencer l'action (motif d'annulation). Le bénéfice central est le **découplage** : le publieur ignore ses abonnés, qui peuvent être ajoutés, retirés ou multiples — fondement du patron Observer et moyen privilégié de faire communiquer objets métier et interface sans couplage fort.

La section suivante aborde deux mécanismes de structuration supplémentaires : les [collections personnalisées et les interfaces (Implements)](05-collections-interfaces.md).

⏭️ [16.5. Collections personnalisées et interfaces (Implements)](/16-poo-vba-access/05-collections-interfaces.md)
