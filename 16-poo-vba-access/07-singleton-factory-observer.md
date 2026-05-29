🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 16.7. Patrons Singleton, Factory et Observer en VBA

Les **patrons de conception** sont des solutions éprouvées à des problèmes récurrents. Trois d'entre eux reviennent constamment dans une application orientée objet, et ont déjà affleuré dans les sections précédentes : la **fabrique** (pour créer des objets, évoquée en 16.1 et 16.6), l'**observateur** (pour notifier, fondé sur les événements de 16.4), et le **singleton** (pour garantir une instance unique). Cette section les présente dans leur adaptation à VBA, dont les contraintes — absence de constructeur paramétré, de membres statiques et d'héritage — façonnent les implémentations.

## Le Singleton

### Intention

Le Singleton garantit qu'une classe ne possède qu'**une seule instance**, accessible globalement. On l'emploie pour les objets dont il ne doit exister qu'un exemplaire partagé dans toute l'application : un objet de configuration, un journaliseur (logger), un cache, un gestionnaire de connexion.

### Mise en œuvre en VBA

VBA n'a ni constructeur privé ni membre statique : on ne peut donc pas *interdire* le `New` au niveau du langage. Le Singleton se réalise par **convention**, au moyen d'une fonction de module standard qui conserve l'instance et la crée au premier appel.

```vba
' --- Module standard : modConfig ---
Option Explicit
Private m_Instance As clsConfig

Public Function Config() As clsConfig
    If m_Instance Is Nothing Then
        Set m_Instance = New clsConfig
        m_Instance.Charger              ' initialisation au premier accès
    End If
    Set Config = m_Instance
End Function
```

Partout dans l'application, on accède alors à l'unique instance via `Config` :

```vba
Debug.Print Config.NomEntreprise
```

La variable `m_Instance`, de niveau module, persiste pour la durée de la session, à la manière d'une variable globale.

### Limites et prudence

Trois réserves accompagnent ce patron. D'abord, VBA ne peut pas **garantir** l'unicité : rien n'empêche techniquement un `New clsConfig` ailleurs ; le Singleton repose sur la discipline (« toujours passer par `Config`, jamais `New` »). Ensuite, comme pour toute variable de module non compilée, l'instance peut être perdue lors d'une réinitialisation provoquée par une erreur non gérée (rappel de la nuance vue en [section 15.7](/15-multi-utilisateurs/07-tempvars.md)). Enfin, et surtout, un Singleton est de l'**état global déguisé** : en abuser crée des dépendances cachées et nuit à la testabilité. On le réserve donc à quelques objets réellement uniques (configuration, journalisation), sans en faire un réflexe.

## La Factory (fabrique)

### Intention

Une fabrique encapsule la **création** d'objets. Plutôt que de disséminer des `New` — surtout lorsque l'initialisation est complexe ou que le choix de l'implémentation dépend du contexte —, on centralise la construction. La fabrique rend deux services distincts : construire un objet correctement initialisé, et choisir quelle classe concrète instancier.

### Fabrique de construction paramétrée

Premier usage : pallier l'absence de constructeur paramétré ([section 16.1](01-modules-classe.md)). Une fonction fabrique crée l'objet et le renvoie déjà configuré, en une seule opération.

```vba
Public Function CreerClient(ByVal id As Long, ByVal nom As String, _
                            ByVal email As String) As clsClient
    Dim c As clsClient
    Set c = New clsClient
    c.Init id, nom, email
    Set CreerClient = c
End Function
```

```vba
Set c = CreerClient(42, "Dupont", "d@exemple.fr")   ' objet prêt à l'emploi
```

Un objet ne peut ainsi jamais exister dans un état non initialisé : on l'obtient toujours complet.

### Fabrique de choix d'implémentation

Second usage, plus puissant : renvoyer un objet du **type d'une interface** en choisissant la classe concrète selon la configuration. C'est exactement ce qui centralise le choix d'implémentation du Repository ([section 16.6](06-patron-repository.md)).

```vba
' --- Module standard : modFabriques ---
Public Function CreerClientRepository() As IClientRepository
    Select Case Config.Moteur           ' paramètre de configuration (cf. Singleton ci-dessus)
        Case "SQLServer"
            Set CreerClientRepository = New clsClientRepositorySqlServer
        Case Else
            Set CreerClientRepository = New clsClientRepositoryDAO
    End Select
End Function
```

```vba
Dim repo As IClientRepository
Set repo = CreerClientRepository()       ' le consommateur ignore l'implémentation
```

Le choix de l'implémentation ne figure plus qu'**ici**. Migrer vers un autre moteur ne touche que cette fabrique — le bénéfice de migration évoqué en 16.6 trouve là son point d'ancrage unique. Cet exemple combine d'ailleurs les trois idées : une fabrique qui interroge un singleton de configuration pour fournir une implémentation de repository.

## L'Observer (observateur)

### Intention

L'Observer définit une dépendance **un-à-plusieurs** : quand un objet (le *sujet*) change d'état, tous ses dépendants (les *observateurs*) en sont notifiés automatiquement. En VBA, ce patron est très naturellement servi par les **événements personnalisés** de la [section 16.4](04-evenements-personnalises-classe.md).

### Observer par événements (la voie idiomatique)

Le sujet déclare et déclenche des événements ; les observateurs déclarent le sujet `WithEvents` et écrivent des gestionnaires. C'est précisément le mécanisme de la section 16.4 : un publieur peut avoir plusieurs abonnés, chacun détenant sa propre référence `WithEvents` — ce qui réalise directement la relation un-à-plusieurs. Pour quelques observateurs et un usage simple, c'est la solution la plus directe et la mieux outillée par l'éditeur.

Ses limites sont celles de `WithEvents` (rappelées en 16.4) : déclaration au niveau module, hôte de type classe ou formulaire, liaison anticipée, un sujet par variable.

### Observer par interface (la voie classique)

Lorsque le nombre d'observateurs est dynamique et géré au moment de l'exécution, on revient au patron tel que décrit dans sa forme classique : une **interface** `IObservateur` et une **collection d'observateurs** maintenue par le sujet, qu'il parcourt pour notifier chacun.

```vba
' --- Module de classe : IObservateur (interface) ---
Option Explicit
Public Sub Actualiser(ByVal source As Object, ByVal info As String)
End Sub
```

```vba
' --- Dans le sujet observé (ex. clsTache) ---
Private m_Observateurs As Collection        ' d'objets IObservateur

Private Sub Class_Initialize()
    Set m_Observateurs = New Collection
End Sub

Public Sub Abonner(ByVal obs As IObservateur)
    m_Observateurs.Add obs
End Sub

Public Sub Desabonner(ByVal index As Long)
    m_Observateurs.Remove index
End Sub

' Le sujet notifie tous ses observateurs lorsqu'il change d'état
Private Sub NotifierTous(ByVal info As String)
    Dim obs As IObservateur
    For Each obs In m_Observateurs
        obs.Actualiser Me, info
    Next obs
End Sub
```

Chaque observateur est une classe qui `Implements IObservateur` et fournit `IObservateur_Actualiser`. Cette forme repose sur les interfaces et les collections de la [section 16.5](05-collections-interfaces.md).

### Événements ou interface ?

Les deux réalisent l'Observer, avec des compromis différents. Les **événements** sont idiomatiques, concis et bien assistés par l'éditeur, mais soumis aux contraintes de `WithEvents`. L'**interface avec collection** est plus verbeuse, mais plus souple lorsqu'il faut gérer dynamiquement un nombre variable d'observateurs de classes diverses, sans la contrainte d'hôte de `WithEvents`. On choisit les événements par défaut, l'interface quand la gestion dynamique l'exige.

## Points de vigilance

- **Singleton : convention, pas garantie.** VBA n'empêche pas un `New` direct ; toujours passer par l'accesseur unique.
- **Singleton avec modération** : c'est de l'état global ; en abuser crée des dépendances cachées et nuit aux tests.
- **Fabrique de construction** : le moyen propre de simuler un constructeur paramétré (16.1).
- **Fabrique de choix** : centraliser à un seul endroit le choix d'implémentation (clé de la migration, 16.6).
- **Observer : privilégier les événements** (16.4) ; recourir à l'interface + collection pour un nombre d'observateurs dynamique.
- **Désabonnement** : prévoir un moyen de retirer un observateur, et libérer les références pour éviter de maintenir des objets en vie inutilement.

## En résumé

Trois patrons classiques s'adaptent aux particularités de VBA. Le **Singleton** garantit une instance unique via une fonction de module standard qui la crée au premier appel et la conserve — une garantie *par convention*, VBA ne pouvant interdire le `New`, et à employer avec parcimonie car il s'agit d'état global. La **Factory** centralise la création d'objets : sous forme de fonction de construction paramétrée, elle pallie l'absence de constructeur à arguments (16.1) ; sous forme de fabrique renvoyant une interface, elle concentre en un seul point le choix de l'implémentation, ce qui sous-tend la souplesse de migration du Repository (16.6). L'**Observer** établit une notification un-à-plusieurs, réalisée naturellement par les **événements** de VBA (16.4), ou, pour un ensemble d'observateurs dynamique, par une **interface `IObservateur` et une collection** (16.5). Ensemble, ces patrons fournissent les briques de structuration qui mènent à l'architecture en couches.

La section suivante assemble l'ensemble du chapitre en une organisation cohérente : l'[architecture en couches (DAL, BLL, UI) appliquée à Access](08-architecture-en-couches.md).

⏭️ [16.8. Architecture en couches (DAL, BLL, UI) appliquée à Access](/16-poo-vba-access/08-architecture-en-couches.md)
