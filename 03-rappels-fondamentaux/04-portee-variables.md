🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 3.4. Portée des variables et durée de vie

Quatrième rappel : la **portée** et la **durée de vie** des variables — deux notions souvent confondues, alors qu'elles répondent à des questions différentes. La portée dit **où** une variable est visible ; la durée de vie dit **combien de temps** elle conserve sa valeur. Bien les distinguer évite des bugs subtils, et certaines particularités d'Access rendent ce sujet d'autant plus important.

## Deux notions à ne pas confondre

- La **portée (scope)** détermine **depuis quels endroits du code** une variable est accessible.
- La **durée de vie (lifetime)** détermine **pendant combien de temps** elle existe et garde sa valeur.

Ces deux propriétés sont liées, mais indépendantes — et le mot-clé `Static`, vu plus loin, en est la preuve : il offre une portée très restreinte tout en allongeant la durée de vie.

## Les trois niveaux de portée

VBA propose trois niveaux, du plus restreint au plus large.

**Portée locale (procédure).** Une variable déclarée par `Dim` **à l'intérieur d'une procédure** n'est visible que dans cette procédure. C'est le niveau à privilégier par défaut.

**Portée module (privée).** Une variable déclarée dans la **section Déclarations** (en tête du module, avant toute procédure) avec `Dim` ou `Private` est visible par **toutes les procédures du module**, mais par aucune autre.

**Portée projet (publique).** Une variable déclarée dans la section Déclarations d'un **module standard** avec `Public` est visible depuis **tous les modules du projet**.

```vba
' --- Section Déclarations du module ---
Public gAppNom As String          ' portée projet
Private mCompteurModule As Long    ' portée module

Sub Exemple()
    Dim x As Long                  ' portée locale
    x = 1
    ' x, mCompteurModule et gAppNom sont tous accessibles ici
End Sub
```

Une **nuance importante**, déjà rencontrée aux sections 2.3 et 3.3 : `Public` n'a ce sens « global » que dans un **module standard**. Dans un module de **classe**, de **formulaire** ou d'**état**, un membre `Public` devient une **propriété de l'objet**, accessible à travers une instance — et non une variable globale flottante.

## La durée de vie

La durée de vie dépend du niveau de déclaration :

- une variable **locale** (`Dim` dans une procédure) est **créée à chaque appel** et **détruite à la fin** de la procédure ; elle ne conserve donc rien d'un appel au suivant ;
- les variables de **module** et de **projet** vivent **aussi longtemps que le projet est chargé**, c'est-à-dire toute la session.

## Le mot-clé Static

`Static` change la donne pour une variable locale : elle reste **invisible hors de la procédure** (portée locale), mais **conserve sa valeur entre les appels** (durée de vie étendue à la session). C'est l'outil idéal pour un compteur ou un état « à mémoire » sans pour autant exposer une variable globale.

```vba
Sub Incrementer()
    Static nbAppels As Long        ' conserve sa valeur entre les appels
    nbAppels = nbAppels + 1
    Debug.Print "Appel n° " & nbAppels
End Sub
' Appels successifs : affiche 1, puis 2, puis 3...
```

Avec un simple `Dim`, le comportement serait tout autre :

```vba
Sub IncrementerDim()
    Dim nbAppels As Long           ' recréée à 0 à chaque appel
    nbAppels = nbAppels + 1
    Debug.Print nbAppels           ' affiche toujours 1
End Sub
```

On peut aussi rendre **toute une procédure** `Static` (`Static Sub …`), ce qui rend statiques l'ensemble de ses variables locales ; usage plus rare.

## Tableau récapitulatif

| Niveau de portée | Où la déclarer | Mot-clé | Visible depuis | Durée de vie |
|---|---|---|---|---|
| **Procédure (locale)** | dans la procédure | `Dim` | la procédure seule | l'exécution de la procédure |
| **Procédure (persistante)** | dans la procédure | `Static` | la procédure seule | toute la session |
| **Module (privée)** | section Déclarations | `Dim` / `Private` | toutes les procédures du module | toute la session |
| **Projet (publique)** | section Déclarations (module standard) | `Public` | tous les modules du projet | toute la session |

## Le piège Access : la perte de l'état global

Voici un point **spécifique au contexte Access** (et à VBA en général) qu'il faut absolument connaître. Les variables de **module** et **publiques** sont censées vivre toute la session… mais une **erreur non gérée** peut **réinitialiser l'ensemble de l'état du projet** : lorsqu'une erreur fatale survient et que l'on choisit « Fin » (ou que l'on arrête le code, ou qu'on le recompile), **toutes ces variables retombent à leur valeur par défaut**.

Conséquence : on ne peut pas se fier aveuglément à une variable globale pour stocker un état **critique** censé persister tout au long de l'application — un utilisateur connecté, un paramètre de session, etc. Le moindre incident risque de l'effacer.

> 💡 L'alternative robuste sous Access s'appelle **`TempVars`** : des variables de session qui **survivent** à ces réinitialisations. Pour tout état devant persister de façon fiable durant la session, elles sont préférables à une variable `Public`. Les `TempVars` sont traitées en section 15.7.

```vba
' Variable globale classique : perdue en cas d'erreur non gérée
Public gUtilisateur As String

' Alternative robuste sous Access (chapitre 15.7)
TempVars!Utilisateur = "Marie"
Debug.Print TempVars!Utilisateur
```

## Bonnes pratiques

- Déclarez chaque variable dans la **portée la plus étroite** suffisante (locale de préférence, puis module, puis projet). Cela réduit le couplage, les bugs et les conflits de noms.
- **Limitez les variables globales** : elles rendent le code plus difficile à raisonner et sont **fragiles** (réinitialisées sur erreur).
- Réservez **`Static`** au besoin précis de « mémoire entre appels ».
- Pour un **état de session** devant persister sous Access, préférez les **`TempVars`** (section 15.7).

## Conflits de noms (masquage)

Lorsqu'une variable locale porte le **même nom** qu'une variable de module ou publique, la locale **masque** l'autre à l'intérieur de la procédure : c'est elle qui l'emporte. Ce comportement, bien que défini, est source de confusion ; mieux vaut **éviter les noms identiques** entre niveaux de portée.

## Le cas des constantes

Les **constantes** suivent exactement les mêmes niveaux de portée que les variables : locale (dans une procédure), `Private` (module) ou `Public` (projet, module standard). C'est ce qui avait été évoqué en section 3.1.

## Variables objet : penser à libérer

Pour les variables d'**objet**, la fin de la portée libère la référence ; on peut aussi la libérer explicitement avec `Set obj = Nothing`. Pour les objets de portée **module** ou **projet** (par exemple un objet DAO conservé longtemps), il est recommandé de les libérer une fois inutiles. Cette hygiène de fermeture est détaillée pour DAO en section 9.11.

```vba
Dim db As DAO.Database
Set db = CurrentDb
' ... utilisation ...
Set db = Nothing        ' libère la référence
```

## À retenir

- **Portée** = *où* une variable est visible ; **durée de vie** = *combien de temps* elle garde sa valeur. Les deux sont distinctes.
- Trois niveaux de portée : **locale** (`Dim` en procédure), **module** (`Dim`/`Private` en Déclarations), **projet** (`Public` en module standard).
- **`Static`** combine portée locale et durée de vie étendue : la variable **conserve sa valeur entre les appels**.
- **Piège Access** : une **erreur non gérée réinitialise** les variables globales ; pour un état de session fiable, utilisez les **`TempVars`** (section 15.7).
- Déclarez au plus **étroit**, limitez les **globales**, et pensez à **libérer** les variables objet de longue durée (`Set … = Nothing`, section 9.11).

---


⏭️ [3.5. Tableaux statiques et dynamiques](/03-rappels-fondamentaux/05-tableaux.md)
