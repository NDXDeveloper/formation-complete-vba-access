🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 19.1. Points d'arrêt, exécution pas à pas et fenêtre Espion

Le débogage interactif est le moyen le plus direct de voir ce que fait réellement le code : on suspend l'exécution, on l'avance instruction par instruction, et l'on inspecte l'état des variables à chaque pas. L'éditeur VBA réunit pour cela tout l'outillage nécessaire — points d'arrêt, exécution pas à pas et fenêtre Espion. C'est une activité de **développement** : elle suppose l'accès à l'éditeur, dont on ne dispose pas sur une application déployée (voir la [section 19.7](07-debogage-distance.md) pour ce cas).

## Le mode arrêt : suspendre l'exécution

VBA connaît trois modes : conception, exécution et **arrêt** (break). C'est en mode arrêt que tout se joue : l'exécution est figée sur une ligne précise, surlignée en jaune, et l'on peut alors examiner les variables, modifier leur valeur, voire évaluer des expressions. On entre en mode arrêt par un point d'arrêt, par l'instruction `Stop`, par une exécution pas à pas, ou à la suite d'une erreur non gérée.

## Les points d'arrêt

Un point d'arrêt suspend l'exécution **juste avant** que la ligne marquée ne s'exécute. On le pose ou on le retire en cliquant dans la marge grise à gauche de la ligne, ou avec la touche **F9** (Activer/Désactiver le point d'arrêt). La ligne se signale alors par une pastille dans la marge.

Deux points méritent l'attention. D'une part, on ne peut pas poser de point d'arrêt sur une ligne non exécutable (déclaration, commentaire, ligne vide). D'autre part — et c'est important — **les points d'arrêt ne sont pas enregistrés** avec le projet : ils disparaissent à la fermeture de la base. Pour effacer tous les points d'arrêt d'un coup, on utilise **Ctrl+Maj+F9** (Effacer tous les points d'arrêt).

## Le mot-clé `Stop`

Lorsqu'on souhaite un arrêt **persistant**, inscrit dans le code, on emploie l'instruction `Stop` : elle provoque un passage en mode arrêt comme un point d'arrêt, mais survit à la fermeture du projet.

```vba
If montant < 0 Then
    Stop          ' arrêt forcé : la valeur ne devrait jamais être négative
End If
```

Cette persistance est aussi un danger : un `Stop` oublié dans une application livrée interromprait l'exécution chez l'utilisateur. On veille donc à retirer tous les `Stop` avant le déploiement.

## Avancer pas à pas

Une fois en mode arrêt, plusieurs commandes permettent de progresser dans le code à la vitesse voulue.

| Commande | Raccourci | Effet |
|---|---|---|
| Pas à pas détaillé (Step Into) | F8 | Exécute la ligne ; **entre** dans la procédure appelée |
| Pas à pas principal (Step Over) | Maj+F8 | Exécute la ligne ; exécute la procédure appelée **sans y entrer** |
| Pas à pas sortant (Step Out) | Ctrl+Maj+F8 | Termine la procédure courante et s'arrête au retour |
| Exécuter jusqu'au curseur | Ctrl+F8 | Reprend jusqu'à la ligne où se trouve le curseur |
| Définir l'instruction suivante | Ctrl+F9 | Déplace le pointeur d'exécution sur une autre ligne |
| Continuer | F5 | Reprend l'exécution normale jusqu'au prochain arrêt |
| Réinitialiser | — | Arrête tout et remet l'état des variables à zéro |

Le **pas à pas principal** (Maj+F8) est précieux pour survoler les appels dont on est sûr, sans s'y enfoncer, tandis que le **pas à pas sortant** permet de remonter rapidement quand on est entré par erreur dans une procédure. La commande **Définir l'instruction suivante** déplace la flèche jaune d'exécution — en la faisant glisser dans la marge ou via le menu — pour sauter du code ou rejouer une ligne ; puissante pour tester une branche, mais à manier avec prudence, car elle peut laisser l'état dans une situation incohérente. Enfin, **Réinitialiser** interrompt totalement l'exécution et efface l'état : les variables de module et `Static` sont alors perdues (voir la [section 18.4](../18-optimisation-performance/04-cache-variables-module.md)).

> ℹ️ L'ensemble des raccourcis de l'éditeur est rassemblé dans l'annexe [I](../annexes/i-raccourcis-clavier-editeur-vba.md).

## Inspecter les valeurs en mode arrêt

Suspendre l'exécution n'a d'intérêt que pour observer l'état du programme. Plusieurs moyens se complètent.

Le plus immédiat est le **survol** : en mode arrêt, poser le pointeur de la souris sur une variable affiche une info-bulle indiquant sa valeur courante.

La **fenêtre Variables locales** offre une vue d'ensemble : elle liste toutes les variables de la procédure en cours avec leur valeur et leur type, et se met à jour à chaque pas. C'est l'outil idéal pour embrasser d'un coup d'œil tout l'état local, plutôt que de survoler les variables une à une.

La **fenêtre Exécution immédiate** permet en outre d'afficher (`?expression`) ou de modifier des valeurs à la volée ; elle fait l'objet de la [section 19.2](02-debug-print-assert.md).

> ℹ️ Les fenêtres de l'éditeur sont présentées au [chapitre 2.4](../02-interface-environnement/04-fenetres-proprietes-execution.md).

## La fenêtre Espion

La fenêtre Espion va plus loin que l'inspection ponctuelle : elle suit en continu des expressions choisies, et peut même déclencher l'arrêt automatiquement selon leur valeur.

### Ajouter un espion

On sélectionne une expression, puis on l'ajoute via le menu Débogage > Ajouter un espion (ou par le menu contextuel). La boîte de dialogue demande l'**expression** à surveiller, son **contexte** (la procédure ou le module où elle s'applique) et le **type** d'espion. Un espion défini pour une procédure précise n'est évalué que dans ce contexte ; ailleurs, il affiche « hors contexte ».

### Les trois types d'espion

C'est là que réside la véritable puissance de l'outil, au-delà du simple point d'arrêt.

L'**expression espionnée** se contente d'afficher la valeur de l'expression, réactualisée à chaque entrée en mode arrêt et à chaque pas.

L'**arrêt si la valeur est vraie** suspend automatiquement l'exécution dès que l'expression devient vraie (non nulle) : c'est un point d'arrêt conditionnel par valeur. Pour s'arrêter au 500ᵉ tour d'une boucle, par exemple, on surveille `i = 500` en mode « arrêt si vrai », sans avoir à avancer manuellement.

L'**arrêt si la valeur change** suspend l'exécution chaque fois que la valeur de l'expression se modifie. C'est l'outil de choix pour découvrir *où* une variable est altérée de façon inattendue : on laisse tourner, et VBA s'arrête précisément à l'endroit du changement.

### L'espion express

Pour une évaluation immédiate, l'**espion express** (Quick Watch, **Maj+F9**) affiche aussitôt la valeur de l'expression sélectionnée dans une fenêtre surgissante, avec la possibilité de l'ajouter comme espion permanent.

## Une méthode de débogage interactif

Ces outils prennent tout leur sens enchaînés selon une démarche. On reproduit le défaut, on pose un point d'arrêt à proximité de la zone suspecte, puis on avance pas à pas en observant les variables (Variables locales, survol, espions). Lorsqu'une variable prend une valeur aberrante sans qu'on sache où, on pose un espion « arrêt si la valeur change » pour localiser la mutation. Pour atteindre rapidement un cas particulier au sein d'une boucle ou d'une condition, on combine « arrêt si la valeur est vraie » et « Exécuter jusqu'au curseur ». Pour un code d'événement de formulaire, on pose le point d'arrêt puis on déclenche l'action correspondante dans l'interface : l'arrêt se produit au moment voulu.

## Limites et précautions

Le débogage interactif a ses contraintes. Les points d'arrêt étant perdus à la fermeture, on recourt à `Stop` pour un arrêt durable — quitte à le retirer avant livraison. Éditer le code en mode arrêt peut provoquer une réinitialisation du projet (« Cette action réinitialise votre projet »), avec perte de l'état des variables de module. Le fait de suspendre l'exécution **altère par ailleurs le minutage** : sans effet sur la logique, mais à proscrire pour mesurer des performances (voir la [section 18.1](../18-optimisation-performance/01-profilage-mesure-performances.md)) ou pour du code sensible au temps et aux événements. Enfin, ce mode opératoire ne s'applique qu'en développement : un fichier ACCDE déployé ne s'ouvre pas dans l'éditeur, d'où le recours à d'autres techniques en production.

> ℹ️ Voir le débogage sur poste utilisateur ([19.7](07-debogage-distance.md)), la compilation en ACCDE ([20.3](../20-securite-protection/03-compilation-accde.md)) et la gestion des erreurs ([chapitre 13](../13-gestion-erreurs/README.md)).

## Points clés à retenir

- Le mode arrêt fige l'exécution sur une ligne et permet d'inspecter — et de modifier — l'état du programme.
- Les points d'arrêt (F9) suspendent avant la ligne marquée mais ne sont pas enregistrés ; `Stop` offre un arrêt persistant, à retirer avant le déploiement.
- L'exécution pas à pas se règle finement : entrer dans un appel (F8), le survoler (Maj+F8), en sortir (Ctrl+Maj+F8), aller jusqu'au curseur (Ctrl+F8) ou repositionner le pointeur (Ctrl+F9).
- On inspecte les valeurs par survol, par la fenêtre Variables locales (vue d'ensemble du contexte) et par l'Exécution immédiate.
- La fenêtre Espion suit des expressions et, surtout, arrête l'exécution « si la valeur est vraie » (point d'arrêt conditionnel) ou « si la valeur change » (pour traquer une mutation inattendue) ; l'espion express évalue à la volée.
- Le débogage altère le minutage et reste une activité de développement, inapplicable telle quelle sur une application déployée.

---


⏭️ [19.2. Debug.Print, Assert et fenêtre Exécution immédiate](/19-debogage-tests/02-debug-print-assert.md)
