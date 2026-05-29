🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 6.3. Formulaires en mode simple, continu et feuille de données

Un formulaire peut s'afficher de plusieurs façons, et ce choix influence non seulement l'expérience utilisateur, mais aussi la **manière de coder**. Cette section présente les trois modes principaux — **simple**, **continu** et **feuille de données** —, quand les employer, et surtout ce qu'ils changent pour le développeur.

## Les trois modes d'affichage

Le mode d'affichage d'un formulaire est gouverné par sa propriété **`DefaultView`**. Trois modes dominent :

- **Mode simple (Single Form)** : un **seul enregistrement à la fois**, occupant tout l'écran. C'est la « fiche » classique.
- **Mode continu (Continuous Forms)** : **plusieurs enregistrements**, la mise en page du formulaire étant **répétée** verticalement pour chacun — comme une liste de petites fiches que l'on fait défiler.
- **Feuille de données (Datasheet)** : une **grille** de type tableur (lignes et colonnes), proche de l'apparence d'Excel.

| Valeur de `DefaultView` | Mode |
|---|---|
| 0 | Formulaire unique (simple) |
| 1 | Formulaires continus |
| 2 | Feuille de données |
| 5 | Formulaire double affichage (split) |

Le **formulaire double affichage** (split form, valeur 5) combine les deux : il présente **simultanément** une vue simple et une feuille de données du même jeu de données, synchronisées. Les valeurs 3 et 4 (tableau/graphique croisé dynamique) sont, elles, obsolètes.

## Quand utiliser quel mode

| Mode | Affiche | Usage typique |
|---|---|---|
| **Simple** | un enregistrement à la fois | **saisie / consultation détaillée** d'une fiche (nombreux champs, onglets, sous-formulaires) |
| **Continu** | plusieurs enregistrements, mise en page répétée | **liste personnalisée** avec contrôles ou boutons par ligne |
| **Feuille de données** | grille de type tableur | **consultation/saisie rapide** ; mode naturel des **sous-formulaires** de lignes |

En pratique : le mode **simple** pour éditer une entité en détail, le mode **continu** pour une liste où l'on souhaite une mise en page riche (au-delà d'une simple grille), et la **feuille de données** pour parcourir beaucoup d'enregistrements de façon compacte — c'est d'ailleurs l'affichage habituel des sous-formulaires affichant des lignes de détail (section 6.4).

## Définir et lire le mode

`DefaultView` se règle surtout **en conception**. À l'ouverture, l'argument `View` de `DoCmd.OpenForm` (section 5.2) peut imposer un mode : `acNormal` ouvre dans le mode par défaut du formulaire, tandis qu'`acFormDS` force la **feuille de données**.

```vba
DoCmd.OpenForm "frmLignes", acFormDS      ' ouvrir en feuille de données

' Lire le mode d'affichage courant
Debug.Print Me.CurrentView                ' ex. 1 = mode Formulaire
```

## Ce que les modes changent pour le code

C'est le point essentiel pour le développeur : selon le mode, le code ne se comporte pas tout à fait de la même façon.

### L'enregistrement courant

En mode **continu** ou **feuille de données**, plusieurs enregistrements sont visibles, mais le code du formulaire opère toujours sur l'**enregistrement courant** — la ligne active. Ainsi, `Me!txtMontant` désigne le champ de l'enregistrement **en cours**, et non « tous » les enregistrements affichés.

```vba
' La ligne active, quel que soit le nombre d'enregistrements visibles
Private Sub btnDetail_Click()
    MsgBox "Montant de la ligne courante : " & Me!txtMontant
End Sub
```

### La section Détail qui se répète : le piège du contrôle partagé

Voici le piège le plus important. En mode continu, la **section Détail est répétée** pour chaque enregistrement — mais un contrôle qui s'y trouve reste **un seul et même objet**, partagé par toutes les lignes ; sa valeur reflète simplement la ligne courante.

Conséquence : on **ne peut pas** modifier l'apparence d'**une seule ligne** par du code (par exemple colorer le fond d'un enregistrement précis), puisque le contrôle est commun à toutes. Pour une apparence **par enregistrement**, il faut recourir à la **mise en forme conditionnelle** (section 17.8), conçue exactement pour cela.

```vba
' ⚠️ En mode continu, un contrôle est PARTAGÉ par toutes les lignes.
'    Me!txtMontant.BackColor = vbRed colorerait TOUTES les lignes, pas une seule.
'    -> pour colorer une ligne selon une condition : mise en forme conditionnelle (17.8)
```

### En-tête/pied de formulaire vs section Détail

Seule la section **Détail** se répète. Les contrôles placés dans l'**en-tête** ou le **pied** de formulaire n'apparaissent **qu'une fois**. C'est pourquoi on y place les éléments « globaux » : un champ de recherche, un total général, des boutons d'action — tandis que les données ligne par ligne vont dans le Détail.

### Les événements

Le mode influence aussi le déclenchement de certains événements : passer d'une ligne à l'autre déclenche l'événement **`Current`** du formulaire (chapitre 8). En feuille de données et en mode continu, naviguer entre enregistrements le fait donc se redéclencher à chaque changement de ligne.

## Les sous-formulaires et les modes

Les modes prennent tout leur sens avec les **sous-formulaires** : un formulaire parent en mode **simple** héberge fréquemment un sous-formulaire en **feuille de données** ou en **continu**, affichant les lignes liées (une commande et ses lignes, par exemple). Cette combinaison est au cœur de la section 6.4.

## À retenir

- Le mode d'affichage est fixé par **`DefaultView`** : **simple** (un enregistrement), **continu** (plusieurs, mise en page répétée), **feuille de données** (grille) — plus le **double affichage** (split).
- **Simple** pour la saisie détaillée, **continu** pour une liste personnalisée, **feuille de données** pour la consultation rapide et les **sous-formulaires** de lignes.
- À l'ouverture, l'argument `View` de `DoCmd.OpenForm` (`acFormDS`…) peut imposer un mode ; `Me.CurrentView` renseigne le mode courant.
- **Point clé pour le code** : en continu/feuille de données, `Me!champ` désigne l'enregistrement **courant** (la ligne active).
- **Piège majeur** : en mode continu, un contrôle est **partagé** par toutes les lignes — pour une apparence **par ligne**, utilisez la **mise en forme conditionnelle** (section 17.8), pas le code.

---


⏭️ [6.4. Sous-formulaires — liaison parent/enfant et propriété LinkMasterFields](/06-formulaires/04-sous-formulaires.md)
