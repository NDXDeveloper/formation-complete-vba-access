🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 19.2. Debug.Print, Assert et fenêtre Exécution immédiate

Là où la section précédente suspendait l'exécution pour l'inspecter, celle-ci propose deux approches complémentaires : **observer sans interrompre**, en traçant l'activité du code, et **interagir directement**, en exécutant des instructions à la main. Au cœur de l'une et de l'autre se trouve la fenêtre Exécution immédiate, secondée par l'objet `Debug`. Cette voie est souvent plus rapide que le pas à pas, et reste indispensable pour les défauts intermittents que l'on ne peut pas figer commodément.

## La fenêtre Exécution immédiate

On l'ouvre par **Ctrl+G** (ou Affichage > Fenêtre Exécution immédiate). Elle joue un double rôle : elle reçoit la sortie de `Debug.Print` et des messages d'exécution, et elle sert de **console interactive** où l'on évalue des expressions et lance des instructions. Elle est disponible dans les trois modes — conception, exécution et arrêt —, ce qui en fait un outil permanent.

## L'utiliser comme console interactive

### Évaluer une expression

Le point d'interrogation `?` est un raccourci pour `Print` : il affiche le résultat d'une expression.

```vba
?2 + 2
?DCount("*", "Clients")
?Format(Now, "yyyy-mm-dd hh:nn:ss")
?Forms!frmClients!txtNom          ' en mode arrêt
```

C'est le moyen le plus rapide de vérifier une valeur, le résultat d'une fonction ou l'état d'un objet.

### Exécuter des instructions

Sans le `?`, la ligne est exécutée comme une instruction : on peut appeler une procédure, modifier une propriété, ou changer la valeur d'une variable en mode arrêt. Le caractère `:` permet d'enchaîner plusieurs instructions, y compris une boucle, sur une seule ligne.

```vba
montant = 100                     ' en mode arrêt : modifie la variable
DoCmd.OpenForm "frmClients"
RecalculerTotaux                  ' lance une procédure pour la tester
For i = 1 To 5: Debug.Print i: Next
```

### Tester sans interface

Cette console évite souvent de construire un déclencheur dans l'interface : on lance directement une procédure ou une fonction pour l'éprouver. En mode arrêt, modifier une variable puis poursuivre permet en outre de tester une hypothèse « et si la valeur était… » sans toucher au code.

La fenêtre a toutefois une **mémoire limitée** : son tampon ne conserve qu'un nombre restreint de lignes, et son contenu disparaît à la fermeture. Elle convient à l'observation ponctuelle, non à la conservation d'un volume de traces — pour cela, on journalise en table ou en fichier (voir la [section 19.5](05-journalisation-evenements.md)).

> ℹ️ Les fenêtres de l'éditeur sont présentées au [chapitre 2.4](../02-interface-environnement/04-fenetres-proprietes-execution.md), les points d'arrêt et le pas à pas à la [section 19.1](01-points-arret-pas-a-pas-espion.md).

## `Debug.Print` : tracer sans interrompre

`Debug.Print` écrit dans la fenêtre Exécution immédiate sans suspendre l'exécution. On l'emploie pour suivre le déroulement : afficher des valeurs, marquer des points de passage, observer la progression d'une boucle — une observation non intrusive, là où le pas à pas serait trop lent ou impossible. On préfixe utilement la trace d'un libellé ou d'un horodatage.

```vba
Debug.Print "ChargerDonnees : " & n & " lignes"
Debug.Print client.ID, client.Nom, client.Solde   ' colonnes (séparateur « , »)
```

Le séparateur d'arguments influe sur la mise en forme : la virgule `,` passe à la zone d'impression suivante (présentation en colonnes), tandis que le point-virgule `;` accole les valeurs. Un `;` en fin d'instruction maintient le curseur sur la même ligne pour le `Print` suivant, ce qui permet de composer une ligne par morceaux.

### En production

Les méthodes de l'objet `Debug` sont **inertes dans une application déployée** : sans éditeur, il n'y a pas de fenêtre Exécution immédiate, et `Debug.Print` n'affiche rien — sans pour autant lever d'erreur. Elles n'ont donc pas besoin d'être retirées avant livraison, à la différence de `Stop`. Mais elles ne servent plus à rien en production : pour diagnostiquer chez l'utilisateur, c'est la journalisation qui prend le relais.

Un point de performance subsiste : même si la sortie est invisible en exécution, **les arguments de `Debug.Print` continuent d'être évalués**. Une expression coûteuse placée dans une trace, au sein d'une boucle serrée, garde donc son coût — ce qui rejoint la mise en garde de la [section 18.1](../18-optimisation-performance/01-profilage-mesure-performances.md) contre les traces dans les zones chronométrées.

### Activer ou désactiver la trace

Pour supprimer entièrement ce coût dans une version de production, on encadre les traces par une **compilation conditionnelle** : la constante désactivée exclut les lignes dès la compilation.

```vba
#Const MODE_TRACE = True

Sub Traiter()
    '  ...
    #If MODE_TRACE Then
        Debug.Print "Traiter : total intermédiaire = " & total
    #End If
End Sub
```

En basculant `MODE_TRACE` à `False` (ou via les arguments de compilation conditionnelle du projet), les `Debug.Print` disparaissent du code compilé, sans coût résiduel.

## `Debug.Assert` : vérifier des invariants

`Debug.Assert` évalue une condition : si elle est **fausse**, l'exécution passe en mode arrêt sur cette ligne ; si elle est vraie, rien ne se passe et l'exécution continue. C'est le moyen d'inscrire dans le code des contrôles de cohérence — des invariants que l'on suppose toujours vrais à un endroit donné — et de s'arrêter précisément là où l'hypothèse est violée.

```vba
Public Function Diviser(ByVal a As Double, ByVal b As Double) As Double
    Debug.Assert b <> 0          ' ne devrait jamais être appelé avec 0
    Diviser = a / b
End Function
```

Comme toutes les méthodes `Debug`, l'assertion est **inerte en exécution déployée** : elle ne déclenche aucun arrêt sans l'éditeur. C'est son grand avantage sur `Stop` : un filet de sécurité de développement qui ne coûte rien en production et n'a pas à être retiré.

### Assertions ou gestion d'erreurs ?

Une distinction est essentielle. Les assertions servent à détecter des **erreurs de programmation** — des bugs, des hypothèses cassées — pendant le développement. Elles ne conviennent **pas** aux conditions qui peuvent légitimement survenir en production : saisie utilisateur erronée, fichier absent, enregistrement introuvable. Ces situations attendues relèvent de la validation et de la gestion des erreurs (voir le [chapitre 13](../13-gestion-erreurs/README.md)) : puisque l'assertion est inerte en production, s'en remettre à elle pour un cas réel laisserait le problème sans traitement.

### `Assert` ou `Stop` ?

On préférera presque toujours `Debug.Assert` à `Stop`. L'assertion est **conditionnelle** (elle n'arrête qu'en cas d'anomalie) et **s'efface d'elle-même** en production ; `Stop` est inconditionnel et persiste — avec le risque d'interrompre l'application chez l'utilisateur.

## Quand préférer la trace au pas à pas

La trace par `Debug.Print` complète le débogage interactif de la [section 19.1](01-points-arret-pas-a-pas-espion.md). On la privilégie pour les défauts **intermittents** difficiles à figer, pour suivre la progression d'une **boucle** sans s'arrêter à chaque tour, pour le code **sensible au minutage** que le pas à pas perturberait, et pour obtenir une **vue d'ensemble** du flux d'exécution. Les assertions, elles, s'ajoutent en continu au code pour faire échouer tôt les hypothèses fausses.

## Points clés à retenir

- La fenêtre Exécution immédiate (Ctrl+G) sert à la fois de sortie de trace et de console interactive ; `?expression` évalue, une instruction sans `?` s'exécute.
- Elle permet de tester une procédure ou de modifier une variable en mode arrêt, mais sa mémoire est limitée et éphémère : pas de journalisation de volume.
- `Debug.Print` trace sans interrompre ; inerte en production (aucun affichage, aucune erreur, aucun retrait nécessaire), mais ses arguments restent évalués — d'où la compilation conditionnelle pour les supprimer.
- `Debug.Assert` arrête l'exécution quand un invariant est faux ; inerte en production, ce qui en fait un filet de sécurité supérieur à `Stop`.
- Les assertions visent les bugs de développement, jamais les conditions attendues en production, qui relèvent de la gestion des erreurs.
- La trace complète le pas à pas : on la préfère pour les bugs intermittents, les boucles, le code sensible au temps et la vue d'ensemble.

---


⏭️ [19.3. Techniques de test des modules de données (DAO/ADO)](/19-debogage-tests/03-tests-modules-donnees.md)
