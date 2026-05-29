🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 16.3. Méthodes et encapsulation

Les [propriétés](02-proprietes-get-let-set.md) ont donné aux objets un état contrôlé ; les **méthodes** vont leur donner des **comportements**. Ensemble, état et comportement réunis dans une même unité, avec un intérieur masqué et un extérieur maîtrisé, forment le principe central de la POO : l'**encapsulation**. Cette section présente d'abord les méthodes, puis formalise l'encapsulation que les sections précédentes mettaient déjà en pratique.

## Les méthodes : les comportements d'un objet

Une **méthode** est une action que l'objet sait accomplir. Techniquement, c'est une procédure publique déclarée dans le module de classe :

- une **`Public Sub`** réalise une action sans renvoyer de valeur ;
- une **`Public Function`** réalise une action *et* renvoie une valeur.

On l'appelle sur l'instance, via `obj.NomMethode(arguments)`. Une méthode opère sur l'**état interne** de l'objet — ses champs privés — et peut accepter des paramètres et retourner un résultat.

```vba
' --- Dans clsClient ---

' Méthode-fonction : calcule une valeur à partir de l'état interne
Public Function CalculerRemise(ByVal montant As Currency) As Currency
    Dim taux As Double
    Select Case m_Categorie
        Case "VIP":    taux = 0.15
        Case "Fidele": taux = 0.1
        Case Else:     taux = 0
    End Select
    CalculerRemise = montant * (1 - taux)
End Function

' Méthode-procédure : effectue une action sur l'état
Public Sub Reinitialiser()
    m_Nom = ""
    m_Prenom = ""
    m_Email = ""
    m_Categorie = "Standard"
End Sub
```

Comme pour toute procédure (cf. [section 3.3](/03-rappels-fondamentaux/03-procedures-fonctions.md)), on privilégie le passage de paramètres en `ByVal` sauf nécessité explicite de modifier la variable de l'appelant.

## Méthode ou propriété ?

Une question revient sans cesse : telle fonctionnalité doit-elle être une propriété ou une méthode ? La règle pratique tient à la nature de ce que l'on expose.

Une **propriété** représente une **caractéristique**, un attribut — un nom ou un adjectif : `c.Nom`, `c.Age`, `c.EstActif`. Sa lecture ne devrait avoir ni effet de bord, ni coût important.

Une **méthode** représente une **action**, un comportement — un verbe : `c.Reinitialiser`, `c.CalculerRemise(montant)`, `c.Envoyer`. Elle *fait* quelque chose, prend souvent des paramètres, et peut modifier l'état.

La frontière est parfois ténue entre une propriété calculée (section 16.2) sans paramètre et une méthode-fonction sans paramètre. Le critère : si la valeur est un **attribut** de l'objet, on en fait une propriété ; si c'est une **action** ou si elle prend des arguments, une méthode. Respecter cette distinction donne une interface lisible, où l'on devine d'emblée ce qui est donnée et ce qui est comportement.

## Membres publics et membres privés

C'est ici que se joue l'essentiel. Dans une classe, on distingue deux catégories de membres :

- les membres **publics** (`Public`) constituent l'**interface** de l'objet — ce que le monde extérieur peut voir et utiliser ;
- les membres **privés** (`Private`) appartiennent à l'**implémentation** — ils ne servent qu'à l'intérieur de la classe et restent invisibles de l'extérieur.

Une méthode privée est une **aide interne** au service des méthodes publiques. Le découpage public/privé n'est pas un détail : c'est le mécanisme même de l'encapsulation.

```vba
' Méthode publique qui s'appuie sur une aide privée
Public Function EstComplet() As Boolean
    EstComplet = ChampsObligatoiresRenseignes()
End Function

' Aide privée : invisible de l'extérieur, au service de la classe
Private Function ChampsObligatoiresRenseignes() As Boolean
    ChampsObligatoiresRenseignes = (Len(m_Nom) > 0) And (InStr(m_Email, "@") > 0)
End Function
```

De l'extérieur, on appelle `c.EstComplet` ; la logique de détail (`ChampsObligatoiresRenseignes`) reste cachée et peut évoluer librement.

> Une mise en garde de conception : on pourrait être tenté de doter une entité comme `clsClient` d'une méthode `Sauvegarder` qui exécute elle-même du code DAO. C'est à éviter, car cela mêle deux responsabilités — modéliser un client *et* accéder à la base. Mieux vaut **déléguer la persistance** à une couche dédiée, le patron Repository ([section 16.6](06-patron-repository.md)), comme on le verra. L'entité décrit *ce qu'est* un client ; le repository sait *comment* le ranger.

## L'encapsulation : définition

L'**encapsulation** recouvre deux idées complémentaires.

La première est le **regroupement** : réunir dans un même objet les données et les comportements qui les concernent. Un `clsClient` rassemble l'état d'un client et tout ce qu'on peut en faire — au lieu de disperser ces données et cette logique dans des variables globales et des procédures éparses.

La seconde est le **masquage de l'information** : garder l'état interne privé et n'exposer qu'une **interface publique délibérée**. L'extérieur manipule l'objet à travers ses propriétés et ses méthodes publiques, sans jamais toucher directement à ses entrailles.

## Pourquoi encapsuler ?

Cette discipline apporte des bénéfices concrets.

Elle **protège les invariants** de l'objet : grâce à la validation dans les setters (section 16.2) et à la logique des méthodes, l'objet garantit lui-même sa cohérence. Du code externe ne peut pas le placer dans un état invalide.

Elle **découple** : les appelants ne dépendent que de l'interface publique, pas de l'implémentation. On peut changer la façon dont l'état est stocké ou calculé — refactoriser l'intérieur — sans rien casser à l'extérieur, tant que l'interface reste stable.

Elle **réduit la complexité perçue** : l'appelant voit une interface restreinte et signifiante, non un amas de détails. Et elle **améliore la maintenabilité** : un bug lié à l'état d'un client se cherche dans `clsClient`, qui en est seul responsable.

## Comment encapsuler en VBA

Les outils sont ceux des sections précédentes, assemblés avec discipline :

- des **champs privés** (`Private m_...`) pour l'état ;
- des **propriétés publiques** (16.2) pour un accès contrôlé à cet état, avec validation ;
- des **méthodes publiques** pour les comportements ;
- des **méthodes privées** pour les aides internes ;
- une surface publique aussi **réduite** que possible.

Un principe guide ces choix : **privé par défaut, public par intention**. On commence par tout déclarer privé, et l'on ne rend public que ce qui doit réellement faire partie de l'interface. L'interface reste ainsi minimale, et l'implémentation libre d'évoluer.

## Encapsulation et responsabilité

Une classe bien encapsulée a une **responsabilité claire** : elle modélise un concept, un seul. `clsClient` s'occupe d'un client — pas de l'envoi d'e-mails, pas de l'accès à la base, pas de l'affichage. Ce qui ne relève pas de sa responsabilité est confié ailleurs : la persistance au repository (16.6), l'orchestration aux formulaires, la structuration générale à une architecture en couches ([section 16.8](08-architecture-en-couches.md)). Cette idée — une classe, une responsabilité — sera approfondie sous l'angle architectural au [chapitre 24](/24-bonnes-pratiques-ressources/02-architecture-applicative.md).

## Points de vigilance

- **Propriété = attribut (nom), méthode = action (verbe)** : respecter cette distinction donne une interface claire.
- **Pas d'effet de bord dans un `Get` de propriété** : si « ça agit », c'est une méthode.
- **Privé par défaut, public par intention** : n'exposer que le strict nécessaire.
- **Les aides internes sont `Private`** : elles font partie de l'implémentation, pas de l'interface.
- **Ne pas mêler les responsabilités** : une entité ne fait pas son propre accès aux données ; déléguer au repository (16.6).
- **Une classe, une responsabilité** : un objet qui « fait tout » est un objet mal encapsulé.

## En résumé

Une **méthode** est un comportement de l'objet, déclaré comme une `Public Sub` (action) ou une `Public Function` (action avec retour), opérant sur l'état interne ; on la distingue d'une propriété par sa nature de verbe plutôt que de nom. La distinction entre membres **publics** (l'interface) et **privés** (l'implémentation, y compris les aides internes) est le mécanisme concret de l'**encapsulation** : réunir données et comportements dans une même unité, tout en masquant les détails derrière une interface délibérée. Encapsuler protège les invariants, découple les appelants de l'implémentation, réduit la complexité et facilite la maintenance. En VBA, on l'obtient par des champs privés, des propriétés et méthodes publiques, des aides privées, et le principe « privé par défaut, public par intention ». Une classe bien encapsulée porte une responsabilité unique, déléguant aux autres couches ce qui ne la concerne pas — au premier rang desquelles la persistance, confiée au patron Repository.

La section suivante dote les objets d'une capacité supplémentaire : émettre et écouter des événements, avec les [événements personnalisés (RaiseEvent / WithEvents)](04-evenements-personnalises-classe.md).

⏭️ [16.4. Événements personnalisés dans une classe (RaiseEvent / WithEvents)](/16-poo-vba-access/04-evenements-personnalises-classe.md)
