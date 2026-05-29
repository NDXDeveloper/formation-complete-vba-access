🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 2.7. Conversion de macros Access en code VBA

Cette dernière section du chapitre referme le fil tendu depuis le chapitre 1. La section 1.3 a comparé macros et VBA et indiqué *quand* préférer l'un à l'autre ; elle mentionnait l'existence d'un outil de conversion. Le voici en pratique : Access sait **transformer automatiquement des macros en code VBA**, ce qui est utile pour faire évoluer une application, consolider son code, ou simplement découvrir l'équivalent VBA d'une action.

## Pourquoi convertir une macro en VBA ?

Les raisons rejoignent les critères de décision de la section 1.3. On convertit une macro lorsqu'elle **atteint ses limites** ou que l'on souhaite bénéficier des atouts de VBA :

- accéder à la **logique complète** (boucles, branchements, recordsets) que la macro ne sait pas exprimer ;
- profiter d'une **gestion d'erreurs** robuste et d'un véritable **débogage** ;
- améliorer la **maintenabilité** et permettre la **gestion de versions** (le VBA s'exporte en texte, contrairement aux macros — voir section 24.4) ;
- **consolider** un projet sur une seule technologie ;
- **apprendre** : comme Access ne possède pas d'enregistreur de macros (section 1.2), la conversion est l'un des moyens de découvrir comment une action s'écrit en VBA.

## Deux situations, deux commandes

Access propose deux commandes distinctes, selon que la macro est **autonome** ou **incorporée** à un objet.

### Convertir une macro autonome

Pour une macro figurant dans le volet de navigation :

1. ouvrir la macro en **mode Création** ;
2. sur le ruban **Outils de macro > Création**, cliquer sur **Convertir les macros en Visual Basic**.

Une boîte de dialogue propose deux options : **ajouter la gestion des erreurs** aux fonctions générées, et **inclure les commentaires** de la macro. Access crée alors un **module standard** nommé « Macro convertie - *NomDeLaMacro* », dans lequel chaque macro (ou sous-macro) devient une **fonction**.

### Convertir les macros incorporées d'un formulaire ou d'un état

Pour les macros **rattachées aux événements** d'un formulaire ou d'un état :

1. ouvrir le formulaire (ou l'état) en **mode Création** ;
2. sur le ruban de création, cliquer sur **Convertir les macros du formulaire en Visual Basic** (ou **…de l'état…** pour un état).

Cette fois, les macros incorporées deviennent des **procédures événementielles** dans le **module de l'objet**, et les propriétés d'événement concernées sont automatiquement reconfigurées sur **[Procédure événementielle]** — le formulaire utilise désormais le code à la place des anciennes macros.

## Ce que produit la conversion

Le résultat dépend du point de départ :

- une **macro autonome** donne un module standard contenant une **fonction par macro/sous-macro** ;
- des **macros incorporées** donnent des **procédures d'événement** dans le module du formulaire ou de l'état.

Dans les deux cas, chaque **action** de macro est traduite en son équivalent VBA — le plus souvent une méthode de l'objet **`DoCmd`** (chapitre 5). Si vous avez coché les options, le code inclut un **gestionnaire d'erreurs** et des **commentaires**.

## Un aperçu de la correspondance

Prenons une petite macro nommée `mAccueil`, composée de deux actions : *OuvrirFormulaire* (`frmClients`) puis *ZoneMessage* (« Bienvenue », type Information). Sa conversion, options activées, produit un code de ce genre :

```vba
'------------------------------------------------------------
' mAccueil
'
'------------------------------------------------------------
Function mAccueil()
On Error GoTo mAccueil_Err

    DoCmd.OpenForm "frmClients", acNormal
    MsgBox "Bienvenue", vbInformation, ""

mAccueil_Exit:
    Exit Function

mAccueil_Err:
    MsgBox Error$
    Resume mAccueil_Exit

End Function
```

Plus généralement, les actions courantes se traduisent ainsi :

| Action de macro | Équivalent VBA généré |
|---|---|
| Ouvrir un formulaire (*OpenForm*) | `DoCmd.OpenForm` |
| Ouvrir un état (*OpenReport*) | `DoCmd.OpenReport` |
| Boîte de message (*MessageBox*) | `MsgBox` |
| Exécuter du SQL (*RunSQL*) | `DoCmd.RunSQL` |
| Définir une valeur (*SetValue*) | affectation (`=`) |
| Exécuter du code (*RunCode*) | appel direct de la fonction |
| Actualiser (*Requery*) | `Me.Requery` / `DoCmd.Requery` |
| Fermer (*Close*) | `DoCmd.Close` |

## Les limites à connaître

La conversion rend service, mais elle a plusieurs limites qu'il faut avoir à l'esprit.

- **Le code généré est fonctionnel, mais peu idiomatique.** Il est verbeux, très centré sur `DoCmd`, avec une gestion d'erreurs générique (`MsgBox Error$`) et un style qui reflète d'anciens usages. Considérez-le comme un **point de départ**, pas comme du code final : il gagne presque toujours à être **refactorisé** vers des méthodes objet plus directes (section 5.8) et une gestion d'erreurs structurée (chapitre 13). La refonte de code existant est traitée en section 24.5.
- **Les macros de données ne sont pas convertibles.** Ce sont des déclencheurs au niveau des tables, sans équivalent VBA (section 1.3) ; aucun outil ne les transforme. Seules les macros **autonomes** et **incorporées** (côté interface) sont concernées.
- **La macro autonome d'origine est conservée.** La conversion crée un nouveau module mais **ne supprime pas** la macro. Les objets qui l'appelaient par son nom (un bouton, par exemple) continuent donc de pointer vers la **macro**, pas vers le nouveau code : il faut **mettre à jour ces références** manuellement.
- **Cas de l'`AutoExec`.** Convertir la macro `AutoExec` produit du code dans un module, mais c'est toujours la **macro** `AutoExec` qui s'exécute au démarrage. Il faut alors repenser le démarrage (par exemple via le pont `ExécuterCode` vu en section 1.3, ou un formulaire de démarrage).

## La conversion comme outil d'apprentissage

Au-delà de la migration, la conversion a une vertu **pédagogique**. Faute d'enregistreur de macros (section 1.2), il n'est pas toujours évident de deviner comment coder telle ou telle action. Créer une petite macro réalisant l'action voulue, puis la convertir, révèle l'**équivalent VBA** — une façon détournée mais efficace de progresser, en particulier sur les nombreuses méthodes de `DoCmd`.

## Démarche recommandée

Pour tirer le meilleur parti de la conversion sans en subir les inconvénients :

1. **convertir** la ou les macros concernées ;
2. **lire et comprendre** le code généré ;
3. le **refactoriser** vers un VBA plus propre (méthodes directes, gestion d'erreurs adaptée) ;
4. **mettre à jour les références** (boutons, autres macros, démarrage) ;
5. **tester** soigneusement le comportement ;
6. **supprimer** la macro d'origine une fois la bascule validée.

## À retenir

- Access convertit automatiquement les macros en VBA via deux commandes : **Convertir les macros en Visual Basic** (macro autonome) et **Convertir les macros du formulaire / de l'état en Visual Basic** (macros incorporées).
- Une macro autonome devient un **module standard** (une **fonction** par macro) ; des macros incorporées deviennent des **procédures événementielles** dans le module de l'objet.
- Les actions sont traduites en VBA, le plus souvent en méthodes **`DoCmd`** ; la gestion d'erreurs et les commentaires sont **optionnels**.
- Le code produit est **fonctionnel mais peu idiomatique** : prévoyez de le **refactoriser** (sections 5.8 et 24.5).
- **Limites** : les **macros de données ne sont pas convertibles**, la **macro autonome d'origine est conservée** (références à mettre à jour), et l'`AutoExec` demande une attention particulière.
- La conversion sert aussi d'**outil d'apprentissage**, en l'absence d'enregistreur de macros (section 1.2).

---


> ✅ **Fin du chapitre 2.** La suite se poursuit avec le chapitre 3 — Rappels fondamentaux de programmation VBA

⏭️ [3. Rappels fondamentaux de programmation VBA](/03-rappels-fondamentaux/README.md)
