🔝 Retour au [Sommaire](/SOMMAIRE.md)

# I. Raccourcis clavier de l'éditeur VBA

Cette annexe récapitule les raccourcis clavier de l'éditeur Visual Basic (VBE). Elle complète le **chapitre 2.1** (accès à l'éditeur) et le **chapitre 19** (débogage). Maîtriser ces combinaisons accélère nettement l'écriture et la mise au point du code.

> ℹ️ Conventions : **Maj** = touche Majuscule (Shift), **Échap** = Échappement (Esc), **Origine/Fin** = Début/Fin (Home/End), **Suppr** = Suppression (Delete). Les termes anglais des menus sont indiqués entre parenthèses quand ils aident au repérage.

## Accès et fenêtres

| Raccourci | Action |
|---|---|
| `Alt+F11` | Basculer entre Access et l'éditeur VBA |
| `F7` | Afficher le code de l'objet sélectionné (View Code) |
| `Ctrl+R` | Explorateur de projets (Project Explorer) |
| `F4` | Fenêtre Propriétés (Properties) |
| `F2` | Explorateur d'objets (Object Browser) |
| `Ctrl+G` | Fenêtre Exécution immédiate (Immediate) |
| `Ctrl+L` | Pile des appels (Call Stack) |
| `Ctrl+Tab` / `Ctrl+F6` | Passer d'une fenêtre ouverte à l'autre |
| `F6` | Basculer entre les volets d'une fenêtre de code fractionnée |

## Exécution et débogage

| Raccourci | Action |
|---|---|
| `F5` | Exécuter la procédure / continuer l'exécution (Run / Continue) |
| `F8` | Pas à pas détaillé — entre dans les procédures appelées (Step Into) |
| `Maj+F8` | Pas à pas principal — exécute la procédure appelée d'un bloc (Step Over) |
| `Ctrl+Maj+F8` | Pas à pas sortant — termine la procédure courante (Step Out) |
| `Ctrl+F8` | Exécuter jusqu'au curseur (Run To Cursor) |
| `F9` | Activer/désactiver un point d'arrêt (Toggle Breakpoint) |
| `Ctrl+Maj+F9` | Effacer tous les points d'arrêt (Clear All Breakpoints) |
| `Ctrl+F9` | Définir l'instruction suivante (Set Next Statement) |
| `Maj+F9` | Espion express (Quick Watch) |
| `Ctrl+Pause` | Interrompre l'exécution en cours (Break) |

## Navigation dans le code

| Raccourci | Action |
|---|---|
| `Maj+F2` | Atteindre la définition de l'élément sous le curseur (Go To Definition) |
| `Ctrl+Maj+F2` | Revenir à la position précédente (Go To Last Position) |
| `Ctrl+Bas` / `Ctrl+Haut` | Procédure suivante / précédente |
| `Ctrl+Page suiv.` / `Ctrl+Page préc.` | Défiler d'un écran vers le bas / le haut |
| `Ctrl+Origine` | Début du module |
| `Ctrl+Fin` | Fin du module |
| `Ctrl+Flèche droite` / `Ctrl+Flèche gauche` | Mot suivant / précédent |

## Recherche

| Raccourci | Action |
|---|---|
| `Ctrl+F` | Rechercher (Find) |
| `Ctrl+H` | Remplacer (Replace) |
| `F3` | Rechercher l'occurrence suivante (Find Next) |
| `Maj+F3` | Rechercher l'occurrence précédente |

## Saisie assistée (IntelliSense)

| Raccourci | Action |
|---|---|
| `Ctrl+Espace` | Compléter le mot en cours (Complete Word) |
| `Ctrl+J` | Liste des propriétés et méthodes (List Properties/Methods) |
| `Ctrl+Maj+J` | Liste des constantes (List Constants) |
| `Ctrl+I` | Informations rapides sur l'élément (Quick Info) |
| `Ctrl+Maj+I` | Informations sur les paramètres (Parameter Info) |

## Édition

| Raccourci | Action |
|---|---|
| `Ctrl+C` / `Ctrl+X` / `Ctrl+V` | Copier / Couper / Coller |
| `Ctrl+Z` | Annuler |
| `Ctrl+Y` | Couper la ligne entière courante (⚠️ **pas** « Rétablir ») |
| `Tab` / `Maj+Tab` | Augmenter / diminuer le retrait des lignes sélectionnées |
| `Ctrl+Suppr` | Supprimer jusqu'à la fin du mot |
| `Ctrl+Retour arrière` | Supprimer jusqu'au début du mot |

## Aide

| Raccourci | Action |
|---|---|
| `F1` | Aide contextuelle sur l'élément sous le curseur |

---

## Points de vigilance

> ⚠️ **`Ctrl+Y` n'est pas « Rétablir ».** Dans l'éditeur VBA, `Ctrl+Y` **coupe la ligne courante** — contrairement à la plupart des applications où il rétablit la dernière action annulée. Erreur de réflexe fréquente.

> 💡 **Ordinateurs portables.** Les touches de fonction (`F5`, `F8`, `F9`…) nécessitent parfois la touche `Fn` selon le réglage du clavier.

> 💡 **Conflits possibles de `Ctrl+Espace`.** Sur certains systèmes, `Ctrl+Espace` est capté par la bascule de langue de saisie (IME). Si la complétion de mot ne réagit pas, vérifier les raccourcis du système d'exploitation.

> 💡 **`F8` est le raccourci le plus utile au débogage.** Combiné aux points d'arrêt (`F9`) et à l'espion express (`Maj+F9`), il constitue le trio de base pour exécuter le code pas à pas et inspecter les valeurs (chapitres 19.1 et 19.2).

---

> 💡 **À retenir.** Quelques raccourcis suffisent au quotidien : **`Alt+F11`** (basculer vers l'éditeur), **`F5`/`F8`/`F9`** (exécuter, pas à pas, point d'arrêt), **`Ctrl+G`** (Exécution immédiate) et **`Maj+F2`** (atteindre la définition). Le reste s'acquiert progressivement. La liste complète figure aussi dans l'aide de l'éditeur (`F1`).

⏭️ [J. Glossaire des termes Access et base de données](/annexes/j-glossaire-access.md)
