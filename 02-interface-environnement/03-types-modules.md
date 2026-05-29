🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 2.3. Types de modules : Standard, Classe, Form, Report

La section 2.2 a montré que l'explorateur de projets regroupe le code en trois dossiers : *Microsoft Access Class Objects*, *Modules* et *Class Modules*. Derrière cette organisation se cachent en réalité **quatre types de modules**, aux rôles bien distincts. Choisir le bon type — et y placer le bon code — est une décision structurante : c'est ce qui distingue une application bien rangée d'un assemblage difficile à maintenir.

## Deux grandes familles

Avant de détailler les quatre types, il faut saisir une distinction de fond. Les modules se répartissent en **deux familles** :

- les **modules standard**, qui sont des **bibliothèques de procédures** — une collection unique de fonctions et de sous-procédures, globalement accessibles, et que l'on n'« instancie » pas ;
- les **modules de classe**, qui sont des **définitions d'objets** — des plans à partir desquels on crée des instances, chacune dotée de ses propres données.

Les **modules de formulaire** et les **modules d'état** ne forment pas une troisième catégorie à part : ce sont des **modules de classe spécialisés**, attachés à un formulaire ou à un état et gérés automatiquement par Access. Garder cette carte mentale — bibliothèque d'un côté, objets de l'autre — éclaire tout le reste.

## Les modules standard

Un **module standard** contient des procédures (`Sub`) et des fonctions (`Function`) à vocation **générale et partagée**. Il n'est rattaché à aucun objet et n'est jamais instancié : il existe simplement comme un réservoir de code accessible depuis toute l'application.

```vba
' Module standard : modUtilitaires
Option Compare Database
Option Explicit

Public Function Initiales(ByVal prenom As String, ByVal nom As String) As String
    Initiales = Left$(prenom, 1) & Left$(nom, 1)
End Function
```

Les membres déclarés **`Public`** d'un module standard sont appelables de partout : depuis d'autres modules, depuis le code des formulaires et des états, mais aussi — c'est une propriété précieuse — depuis les **requêtes** et les **expressions**. La fonction ci-dessus peut ainsi être utilisée directement dans une requête :

```sql
SELECT Initiales([Prenom], [Nom]) AS Init FROM tblClients;
```

C'est donc l'endroit naturel pour la **logique réutilisable** : fonctions utilitaires, règles métier partagées, fonctions d'accès aux données, constantes et variables globales. On crée un module standard via **Insert > Module** dans l'éditeur, ou par le ruban **Créer > Module**.

## Les modules de classe

Un **module de classe** définit un **objet personnalisé** : un type doté de ses propres propriétés (données) et méthodes (comportements). Contrairement au module standard, on en crée des **instances**, et l'on peut en avoir plusieurs simultanément.

```vba
' Module de classe : clsClient
Option Compare Database
Option Explicit

Private m_nom As String

Public Property Get Nom() As String
    Nom = m_nom
End Property

Public Property Let Nom(ByVal valeur As String)
    m_nom = valeur
End Property
```

On l'utilise alors comme un objet :

```vba
Dim c As clsClient
Set c = New clsClient
c.Nom = "Dupont"
```

Les modules de classe permettent l'**encapsulation**, les propriétés (`Property Get / Let / Set`), les méthodes, les événements personnalisés, ainsi que des procédures d'initialisation et de destruction (`Class_Initialize`, `Class_Terminate`). C'est le socle de la programmation orientée objet, à laquelle le **chapitre 16** est entièrement consacré ; nous nous limitons ici à les situer. On les crée via **Insert > Module de classe**, ou par le ruban **Créer > Module de classe**.

## Les modules de formulaire et d'état

Les **modules de formulaire** et d'**état** hébergent le code propre à un objet précis — essentiellement ses **gestionnaires d'événements**. Ils portent un nom **imposé** (`Form_frmClients`, `Report_rptFactures`), apparaissent dans *Microsoft Access Class Objects*, et ne sont créés que lorsque l'objet **contient un module** (propriété `HasModule`, voir section 2.2).

```vba
' Module du formulaire : Form_frmClients
Option Compare Database
Option Explicit

Private Sub Form_Current()
    Me.lblInfo.Caption = "Client : " & Nz(Me.txtNom, "(nouveau)")
End Sub
```

Étant des **modules de classe**, ils sont **instanciés** : un formulaire est, en réalité, un objet créé à son ouverture, et le mot-clé **`Me`** y désigne cette instance. C'est d'ailleurs ce qui rend possible — sujet avancé — l'ouverture de **plusieurs instances** d'un même formulaire. Mais leur usage courant reste simple : on y place tout ce qui concerne **ce** formulaire ou **cet** état, et rien d'autre. La gestion des événements est détaillée aux chapitres 6, 7 et 8.

## Différence clé : l'appel depuis une requête

Voici un point pratique souvent méconnu, qui découle directement des familles décrites plus haut. Une fonction **`Public`** placée dans un **module standard** est appelable depuis une requête ou une expression. La **même** fonction placée dans un module de **formulaire**, d'**état** ou de **classe** ne l'est **pas** : elle y devient une *méthode* de l'objet, accessible seulement à travers une instance, et le moteur de requêtes ne sait pas l'invoquer.

La conséquence est concrète : si vous écrivez une fonction destinée à être utilisée dans une requête, dans un champ calculé ou dans une source de contrôle, elle **doit** résider dans un module standard. C'est une erreur fréquente que de la placer derrière un formulaire et de s'étonner ensuite que la requête échoue.

## Où placer quel code ?

| Type de code | Module recommandé |
|---|---|
| Fonction ou procédure réutilisée par plusieurs objets | Module **standard** |
| Fonction à appeler depuis une requête ou une expression | Module **standard** (`Public`) |
| Code propre à un seul formulaire ou état | Module de **ce** formulaire / état |
| Objet personnalisé (données + comportement encapsulés) | Module de **classe** |
| Constantes et variables globales | Module **standard** (section Déclarations, `Public`) |

La règle générale tient en une phrase : **du code partagé dans un module standard, du code spécifique à un objet dans son propre module, des objets métier dans des modules de classe.**

## La section Déclarations et les instructions Option

Quel que soit son type, tout module commence par une **section Déclarations**, située avant la première procédure. On y trouve :

- les instructions **`Option`** : `Option Compare Database` (les comparaisons de chaînes suivent l'ordre de tri de la base, c'est le réglage par défaut sous Access) et `Option Explicit` (déclaration obligatoire des variables, voir section 1.4) ; plus rarement `Option Base` ou `Option Private Module` ;
- les **variables et constantes de niveau module** (leur portée est traitée au chapitre 3) ;
- les **déclarations d'API Windows** (`Declare`), abordées au chapitre 22.

Veiller à ce que chaque module débute par `Option Explicit` reste l'une des meilleures habitudes à prendre, comme on l'a vu en configurant l'environnement.

## Comment créer chaque type

- **Module standard** et **module de classe** : par le menu **Insert** de l'éditeur (*Module* / *Module de classe*), ou par le ruban **Créer** d'Access.
- **Module de formulaire / d'état** : on ne le crée pas explicitement. Il naît **automatiquement** dès que l'on ajoute du code à un formulaire ou à un état (via *Afficher le code* ou le *Générateur de code*, section 2.1), ce qui passe la propriété `HasModule` à Oui.

## Tableau récapitulatif

| Module | Rôle | Instancié ? | Nommage | Création |
|---|---|---|---|---|
| **Standard** | Procédures / fonctions générales partagées | Non (collection unique) | Libre | Insert > Module / ruban Créer |
| **Classe** | Définit un objet personnalisé | Oui (`New`, plusieurs instances) | Libre | Insert > Module de classe / ruban Créer |
| **Formulaire** | Code et événements d'un formulaire | Oui (à l'ouverture du formulaire) | `Form_…` (imposé) | Automatique (`HasModule`) |
| **État** | Code et événements d'un état | Oui (à l'ouverture de l'état) | `Report_…` (imposé) | Automatique (`HasModule`) |

## À retenir

- Il existe **quatre types de modules**, répartis en deux familles : les **modules standard** (bibliothèques de procédures, non instanciées) et les **modules de classe** (définitions d'objets, instanciées).
- Les modules de **formulaire** et d'**état** sont des **modules de classe spécialisés**, attachés à un objet, nommés automatiquement (`Form_…`, `Report_…`) et créés via `HasModule`.
- Placez le **code partagé** dans un module standard, le **code spécifique** à un objet dans son propre module, et les **objets métier** dans des modules de classe.
- **Point crucial** : seules les fonctions `Public` d'un **module standard** sont appelables depuis une **requête** ou une expression.
- Tout module débute par une **section Déclarations** (instructions `Option`, variables de module, déclarations d'API) ; veillez à y maintenir `Option Explicit`.

---


⏭️ [2.4. Fenêtre Propriétés et fenêtre Exécution immédiate](/02-interface-environnement/04-fenetres-proprietes-execution.md)
