🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 8.6. Propriété ControlSource — contrôles liés et indépendants

La propriété **`ControlSource`** détermine la nature profonde d'un contrôle : est-il **lié** à un champ de données, **calculé** à partir d'une expression, ou **indépendant** (sans rattachement) ? Cette distinction conditionne tout : la lecture, l'écriture, l'enregistrement automatique, et même les erreurs que l'on peut rencontrer. C'est l'un des concepts les plus fondamentaux du travail sur les contrôles.

Cette section consolide ce qui a été évoqué au fil du chapitre et précise les trois états possibles, leurs comportements et leurs pièges — notamment l'épineuse référence circulaire.

---

## 8.6.1. Trois états selon ControlSource

| `ControlSource` | État du contrôle | Comportement |
|---|---|---|
| Un **nom de champ** | **Lié** (bound) | Lecture **et** écriture automatiques du champ |
| Une **expression** (`=…`) | **Calculé** | Affiche le résultat, en **lecture seule** |
| **Vide** | **Indépendant** (unbound) | Valeur en mémoire, gérée par code |

Comprendre dans lequel de ces trois états se trouve un contrôle est la première chose à vérifier face à un comportement inattendu.

---

## 8.6.2. Les contrôles liés (bound)

Un contrôle dont `ControlSource` vaut un **nom de champ** est **lié** : Access affiche automatiquement la valeur du champ pour l'enregistrement courant et **réécrit** dans ce champ toute modification de l'utilisateur. La liaison est **bidirectionnelle**.

```
' En conception : ControlSource = Nom
```

C'est le mode par défaut des formulaires de saisie. Il suppose que le formulaire possède un `RecordSource` contenant ce champ (chapitre 6.5). La `Value` du contrôle reflète alors le champ de l'enregistrement affiché.

---

## 8.6.3. Les contrôles calculés

Un contrôle dont `ControlSource` commence par `=` est **calculé** : il affiche le **résultat d'une expression**. Point essentiel : un contrôle calculé est en **lecture seule** — il n'a aucun champ où écrire, on ne peut donc rien y saisir.

```
=[Quantite]*[PrixUnitaire]
=[Prenom] & " " & [Nom]
=Somme([Montant])            ' dans un pied de section (chapitre 7.5)
```

Le résultat se recalcule automatiquement lorsque ses dépendances changent. L'expression peut référencer d'autres champs ou contrôles.

---

## 8.6.4. Les contrôles indépendants (unbound)

Un contrôle dont `ControlSource` est **vide** est **indépendant** : sa `Value` n'est qu'une donnée **en mémoire** le temps de la session, sans rattachement à un champ. Le développeur la lit et l'écrit par code.

Les contrôles indépendants servent à de multiples usages :

- **critères de recherche** sur un formulaire de filtre ;
- **saisies temporaires** ne devant pas être enregistrées telles quelles ;
- contrôles d'un **formulaire indépendant** entièrement piloté par code (chapitre 6.8) ;
- éléments d'**interface** par nature non liés (boutons, étiquettes).

```vba
' Lire/écrire un contrôle indépendant par code
Me.txtCritere.Value = "Paris"
Dim s As String
s = Nz(Me.txtCritere.Value, "")
```

La valeur d'un contrôle indépendant est **perdue** à la fermeture du formulaire : rien n'est persisté.

---

## 8.6.5. Définir ControlSource par code

La propriété se modifie par programmation pour faire passer un contrôle d'un état à l'autre :

```vba
Me.txtNom.ControlSource = "Nom"             ' lier au champ Nom
Me.txtTotal.ControlSource = "=[Qte]*[PU]"   ' rendre calculé (lecture seule)
Me.txtCritere.ControlSource = ""            ' rendre indépendant
```

Modifier `ControlSource` à l'exécution est possible, mais suppose la cohérence avec la source du formulaire (section suivante).

---

## 8.6.6. ControlSource et RecordSource du formulaire

Un contrôle **lié** ne peut afficher un champ que si ce champ figure dans le **`RecordSource`** du formulaire (chapitre 6.5). Dans le cas contraire, le contrôle affiche une **erreur** : le champ est introuvable.

```
RecordSource du formulaire : SELECT ClientID, Nom, Ville FROM T_Clients
   → un contrôle lié à "Nom" ou "Ville" fonctionne
   → un contrôle lié à "Email" (absent) affiche #Nom?
```

La règle est donc : **tout champ lié à un contrôle doit être présent dans la source du formulaire**.

---

## 8.6.7. Lire la valeur : Value dans tous les cas

Quelle que soit la nature du contrôle, sa propriété **`Value`** renvoie la valeur **actuellement affichée** :

- pour un contrôle **lié** : la valeur du champ de l'enregistrement courant ;
- pour un contrôle **calculé** : le résultat de l'expression (en lecture seule) ;
- pour un contrôle **indépendant** : la valeur qu'on lui a affectée.

L'accès en lecture est donc uniforme ; c'est l'**écriture** qui diffère (impossible sur un calculé, automatique vers le champ sur un lié, manuelle sur un indépendant).

---

## 8.6.8. Les erreurs #Nom? et #Erreur

Deux messages d'erreur signalent un problème de source, et il importe de les distinguer :

| Erreur | Cause typique |
|---|---|
| **`#Nom?`** | Un **nom** ne peut être résolu : champ ou contrôle mal orthographié, champ absent du `RecordSource`, fonction inconnue |
| **`#Erreur`** | L'**expression** échoue à l'évaluation : référence circulaire, division par zéro, incompatibilité de type |

`#Nom?` pointe un identifiant introuvable ; `#Erreur` un calcul qui plante. Cette distinction oriente immédiatement le diagnostic.

---

## 8.6.9. Le piège de la référence circulaire (nom = champ)

Le piège le plus subtil survient lorsqu'un contrôle **calculé** porte le **même nom** qu'un champ qu'il référence dans son expression. Access interprète alors le nom comme le **contrôle lui-même**, créant une **référence circulaire** :

```
' MAUVAIS : contrôle nommé "Montant", source "=[Montant]*1.2"
'           → [Montant] désigne le contrôle → référence circulaire → #Erreur

' BON : contrôle nommé "txtMontantTTC", source "=[Montant]*1.2"
'       → [Montant] désigne sans ambiguïté le champ
```

> **Bonne pratique** : nommer les contrôles avec un **préfixe** distinct du champ (`txtMontant` pour le champ `Montant`). Cela élimine d'emblée les références circulaires et clarifie le code. Les conventions de nommage sont traitées au chapitre 24.1 et à l'annexe G (préfixes Leszynski/Reddick).
>
> À noter : un contrôle **lié** portant le même nom que son champ ne pose **aucun** problème (c'est le comportement normal). Seuls les contrôles **calculés** référençant leur propre nom sont concernés.

---

## 8.6.10. Quel état choisir ?

| Besoin | État recommandé |
|---|---|
| Saisir/afficher une donnée rattachée à un enregistrement | **Lié** |
| Afficher une valeur dérivée (total, concaténation) non stockée | **Calculé** |
| Critère de recherche, saisie temporaire, élément d'interface | **Indépendant** |
| Contrôle sur un formulaire entièrement piloté par code | **Indépendant** (chapitre 6.8) |

En règle générale, une valeur **dérivable** d'autres champs se présente en contrôle **calculé** plutôt que d'être stockée (principe de normalisation) — sauf besoin avéré de la conserver (historique, performance).

---

## 8.6.11. Pièges et bonnes pratiques

- **Identifier l'état du contrôle** (lié, calculé, indépendant) en cas de comportement inattendu : tout en découle.
- **Un contrôle calculé est en lecture seule** : ne pas s'attendre à pouvoir y saisir.
- **Le champ d'un contrôle lié doit figurer dans le `RecordSource`** du formulaire, sous peine de `#Nom?` (chapitre 6.5).
- **Distinguer `#Nom?` (nom introuvable) et `#Erreur` (calcul en échec)** pour diagnostiquer rapidement.
- **Préfixer les noms de contrôles** (distincts des champs) pour éviter les références circulaires sur les contrôles calculés (chapitres 24.1 et annexe G).
- **Ne pas stocker une valeur dérivable** sans raison : la présenter en calculé.
- **Les valeurs des contrôles indépendants ne sont pas persistées** : les enregistrer explicitement si nécessaire (chapitre 6.8).

---

## 8.6.12. Récapitulatif

- `ControlSource` définit l'état d'un contrôle : **lié** (nom de champ, lecture/écriture automatiques), **calculé** (expression `=…`, **lecture seule**) ou **indépendant** (vide, valeur en mémoire gérée par code).
- Un contrôle **lié** suppose que son champ figure dans le **`RecordSource`** du formulaire (chapitre 6.5) ; sinon il affiche `#Nom?`.
- La propriété **`Value`** se lit de façon uniforme dans les trois cas ; seule l'**écriture** diffère.
- Deux erreurs à distinguer : **`#Nom?`** (identifiant introuvable) et **`#Erreur`** (expression en échec, dont la **référence circulaire**).
- La référence circulaire frappe les contrôles **calculés** portant le même nom qu'un champ référencé : on l'évite en **préfixant** les noms de contrôles (chapitres 24.1 et annexe G).
- Choisir l'état selon le besoin : **lié** pour les données d'un enregistrement, **calculé** pour les valeurs dérivées non stockées, **indépendant** pour les critères, saisies temporaires et formulaires pilotés par code (chapitre 6.8).

⏭️ [8.7. Cycle de vie complet d'un formulaire (ordre des événements)](/08-controles-evenements/07-cycle-de-vie-formulaire.md)
