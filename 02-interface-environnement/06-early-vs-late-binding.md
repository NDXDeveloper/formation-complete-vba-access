🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 2.6. Liaisons précoces (Early Binding) vs liaisons tardives (Late Binding)

La section précédente s'est achevée sur une question laissée ouverte : faut-il toujours activer une référence ? La réponse tient dans le choix entre deux modes de **liaison** — la liaison **précoce** (*early binding*) et la liaison **tardive** (*late binding*). Ce choix détermine si une référence est nécessaire, et il met en balance le confort de développement d'un côté, la robustesse au déploiement de l'autre. C'est l'un des arbitrages les plus utiles à maîtriser.

## Qu'est-ce que la liaison ?

La « liaison » désigne le moment où VBA **relie** une variable au type d'objet réel qu'elle représente — autrement dit, le moment où il détermine quelles propriétés et méthodes cet objet possède. Deux moments sont possibles :

- **à la compilation** : on déclare un type précis, et VBA connaît d'avance les membres de l'objet. C'est la **liaison précoce**.
- **à l'exécution** : on déclare un type générique, et VBA ne découvre les membres qu'au moment de l'appel. C'est la **liaison tardive**.

Tout découle de cette différence de timing : disponibilité de l'aide à la saisie, détection des erreurs, performances, et dépendance ou non à une référence.

## La liaison précoce

La liaison précoce **nécessite une référence** à la bibliothèque (Outils > Références, section 2.5). On déclare la variable avec son **type spécifique**, et on l'instancie avec `New` (ou un `CreateObject` typé) :

```vba
' Liaison précoce — nécessite une référence à « Microsoft Excel XX.0 Object Library »
Dim xl As Excel.Application
Set xl = New Excel.Application
xl.Visible = True

Dim wb As Excel.Workbook
Set wb = xl.Workbooks.Add
wb.Sheets(1).Range("A1").Value = "Bonjour"
```

Ses **avantages** sont décisifs pendant l'écriture du code :

- l'**aide à la saisie** (liste automatique des membres) — un gain de productivité considérable ;
- le **contrôle à la compilation** — une faute de frappe dans un nom de propriété est détectée avant l'exécution ;
- les **constantes nommées** de la bibliothèque sont disponibles (`xlUp`, `adOpenStatic`…) ;
- une **exécution plus rapide**, car les appels sont résolus à l'avance.

Son **inconvénient** est celui décrit en section 2.5 : la dépendance à la référence, donc une **sensibilité aux versions** et le risque de référence **MANQUANTE** au déploiement.

## La liaison tardive

La liaison tardive ne requiert **aucune référence**. On déclare la variable comme un **`Object`** générique, et on la crée avec **`CreateObject`**, en passant l'identifiant de programme (*ProgID*) sous forme de chaîne :

```vba
' Liaison tardive — aucune référence requise
Dim xl As Object
Set xl = CreateObject("Excel.Application")
xl.Visible = True

Dim wb As Object
Set wb = xl.Workbooks.Add
wb.Sheets(1).Range("A1").Value = "Bonjour"
```

Ses **avantages** sont symétriques de ses inconvénients précédents :

- **aucune dépendance à une référence** — donc pas de problème de version ni de référence manquante ;
- une bien meilleure **robustesse au déploiement**, notamment pour piloter une application dont on ne maîtrise pas la version sur les postes cibles.

Ses **inconvénients** se paient surtout au moment d'écrire :

- **pas d'aide à la saisie** — on tape les noms de membres « à l'aveugle » ;
- **pas de contrôle à la compilation** — les fautes de frappe ne se révèlent qu'à l'exécution ;
- **pas de constantes nommées** de la bibliothèque (voir plus bas) ;
- une **exécution plus lente**, chaque appel étant résolu à l'exécution.

## Comparaison côte à côte

Les deux exemples ci-dessus sont volontairement identiques, à deux différences près : la **déclaration** (`Excel.Application` *vs* `Object`) et la **création** (`New` *vs* `CreateObject`). Le reste du code — l'utilisation des objets — est **rigoureusement le même**. C'est ce qui rend la conversion de l'un vers l'autre assez mécanique, comme on le verra plus loin.

## Le problème des constantes nommées

Voici l'irritant le plus concret de la liaison tardive. Les **énumérations** d'une bibliothèque (ses constantes nommées) ne sont disponibles qu'en liaison précoce. En liaison tardive, une constante comme `xlUp` est **inconnue** ; il faut soit employer sa **valeur littérale**, soit la **redéclarer** soi-même :

```vba
' Liaison précoce : la constante de la bibliothèque est connue
derniereLigne = xl.Cells(xl.Rows.Count, 1).End(xlUp).Row

' Liaison tardive : xlUp est inconnu — on utilise la valeur littérale...
derniereLigne = xl.Cells(xl.Rows.Count, 1).End(-4162).Row   ' -4162 = xlUp

' ... ou, plus lisible, on redéclare la constante
Const xlUp As Long = -4162
derniereLigne = xl.Cells(xl.Rows.Count, 1).End(xlUp).Row
```

Une précision importante : ce sont les constantes de la **bibliothèque externe** qui disparaissent. Les constantes propres à **VBA** (`vbYes`, `vbCrLf`, `vbCancel`…) restent, elles, toujours disponibles, puisqu'elles appartiennent à une bibliothèque référencée en permanence.

## CreateObject et GetObject

Pour créer ou obtenir un objet, deux fonctions existent :

- **`CreateObject("ProgID")`** crée une **nouvelle instance** du composant ;
- **`GetObject`** permet de se **rattacher à une instance déjà en cours d'exécution**, ou d'ouvrir un fichier en tant qu'objet.

`CreateObject` est indispensable en liaison tardive (faute de type à instancier avec `New`) ; il fonctionne aussi en liaison précoce. Ces deux fonctions et l'automation des applications Office sont approfondies au chapitre 22.

## Quand choisir quoi

Le bon arbitrage dépend du contexte, et il existe une approche pragmatique largement adoptée.

- **Développer en liaison précoce.** Pendant l'écriture, l'aide à la saisie et le contrôle à la compilation sont trop précieux pour s'en priver : on ajoute la référence et on déclare les types précis.
- **Déployer en liaison tardive.** Pour piloter des applications externes (Excel, Word, Outlook) dont la **version varie** d'un poste à l'autre, la liaison tardive évite les ruptures dues à une référence manquante.
- **L'approche hybride.** Beaucoup de développeurs **développent en liaison précoce**, puis, avant la livraison, **basculent en liaison tardive** : remplacer les types par `Object`, `New` par `CreateObject`, retirer la référence, et déclarer les quelques constantes nécessaires. On cumule ainsi confort d'écriture et robustesse de déploiement.
- **Cas des objets internes (DAO, Access).** Pour l'accès aux données local, la liaison précoce avec **DAO** est la norme : DAO est référencé par défaut et fourni avec Access, sans risque de décalage de version au sein d'une même installation. Inutile, ici, de recourir à la liaison tardive.

Un mot sur les **performances** : la liaison précoce est plus rapide par appel, mais l'écart est le plus souvent **négligeable** dans une application Access ordinaire. Il ne devient sensible que dans des **boucles serrées** effectuant un très grand nombre d'appels à des objets externes — situation où il mérite alors considération.

## Tableau comparatif

| Critère | Liaison précoce | Liaison tardive |
|---|---|---|
| Référence requise | Oui | Non |
| Déclaration | Type spécifique (`Excel.Application`) | `Object` |
| Création | `New` ou `CreateObject` | `CreateObject` |
| Aide à la saisie | Oui | Non |
| Contrôle à la compilation | Oui | Non (erreurs à l'exécution) |
| Constantes nommées de la bibliothèque | Disponibles | Indisponibles (valeurs en dur) |
| Performance | Meilleure (souvent négligeable) | Moindre |
| Robustesse au déploiement | Sensible aux versions | Robuste |

## À retenir

- La **liaison** détermine quand VBA relie une variable à son type : à la **compilation** (précoce) ou à l'**exécution** (tardive).
- La **liaison précoce** exige une **référence**, déclare un **type précis**, et offre **aide à la saisie**, **contrôle à la compilation** et **constantes nommées** — au prix d'une sensibilité aux versions.
- La **liaison tardive** se passe de référence (`As Object` + `CreateObject`), au prix de l'aide à la saisie, du contrôle à la compilation et des constantes de la bibliothèque — mais avec une **robustesse** supérieure au déploiement.
- En liaison tardive, **redéclarez** les constantes utiles (`Const xlUp As Long = -4162`) ; les constantes **VBA** (`vbYes`…), elles, restent toujours disponibles.
- Approche recommandée : **développer en liaison précoce**, **déployer en liaison tardive** ; pour **DAO** et les objets internes, la liaison précoce reste la norme.

---


⏭️ [2.7. Conversion de macros Access en code VBA](/02-interface-environnement/07-conversion-macros-vba.md)
