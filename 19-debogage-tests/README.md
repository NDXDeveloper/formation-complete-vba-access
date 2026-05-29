🔝 Retour au [Sommaire](/SOMMAIRE.md)

# Chapitre 19 — Débogage et tests

Tout logiciel comporte des défauts ; la vraie question n'est pas de savoir s'il y en aura, mais avec quelle rapidité on les trouve et les corrige — et combien on en intercepte avant qu'ils n'atteignent les utilisateurs. Les applications Access sont souvent construites sans filet : beaucoup de développeurs s'en remettent au seul tâtonnement, alors même que l'éditeur VBA offre des outils de débogage puissants, et qu'une véritable discipline de test reste possible bien que VBA ne fournisse aucun framework natif.

Ce chapitre couvre ces deux activités complémentaires. Le **débogage** est réactif : il s'agit de localiser et de corriger un défaut une fois qu'un comportement anormal se manifeste. Le **test** est proactif : il vise à vérifier le comportement attendu, idéalement avant que le bug ne survienne. L'un et l'autre reposent sur un même principe — rendre le comportement du programme **observable** plutôt que le deviner : observer, reproduire, vérifier.

## Pourquoi ce chapitre

On retrouve ici l'esprit du chapitre précédent sur la performance : de même qu'on n'optimise pas sans mesurer, on ne corrige pas sans comprendre. Modifier du code « au jugé » jusqu'à ce que le symptôme disparaisse ne fait souvent que déplacer le problème. La démarche saine consiste à reproduire le défaut de façon fiable, à l'isoler, à formuler une hypothèse, à la vérifier par des éléments concrets, puis à confirmer la correction.

Le test, lui, transforme cette vérification en filet permanent. Disposer de tests — même légers — permet de modifier et de faire évoluer une application sans craindre de casser ce qui fonctionnait, et de détecter une régression au plus tôt plutôt que de la découvrir en production.

## Les spécificités d'Access et de VBA

Le débogage en VBA présente quelques particularités qui structurent ce chapitre.

Pendant le développement, le débogage est **interactif** : on suspend l'exécution, on l'avance pas à pas, on inspecte les variables dans l'éditeur. Mais une application déployée — un fichier ACCDE sur le poste d'un utilisateur — ne peut pas être ouverte dans l'éditeur VBA. Diagnostiquer un problème qui ne survient que chez l'utilisateur impose alors d'autres moyens : la **journalisation** et des techniques de débogage à distance.

Les applications Access sont par ailleurs **centrées sur les données**. Les défauts mettent fréquemment en jeu l'état de la base, ce qui rend essentiels le test des modules d'accès aux données et l'usage de jeux de données et d'environnements de test maîtrisés — pour ne jamais déboguer directement sur les données de production.

Enfin, l'absence de framework de test intégré conduit à en **construire un soi-même**, léger mais suffisant pour automatiser les vérifications.

## Débogage et gestion des erreurs

Déboguer et gérer les erreurs sont deux faces d'une même préoccupation. Une application bien instrumentée ne laisse pas les erreurs échouer en silence : elle les rend visibles, les journalise et facilite leur reproduction. La gestion des erreurs proprement dite — structures `On Error`, objet `Err`, journalisation en table — fait l'objet du chapitre 13 ; le présent chapitre s'appuie sur ces fondations pour observer et diagnostiquer.

> ℹ️ Voir le [chapitre 13 — Gestion des erreurs](../13-gestion-erreurs/README.md), en particulier la journalisation des erreurs en table ([13.6](../13-gestion-erreurs/06-journalisation-erreurs-table.md)).

## Ce que vous allez apprendre

Le chapitre progresse des outils interactifs de l'éditeur vers les techniques applicables en production :

- **[19.1. Points d'arrêt, exécution pas à pas et fenêtre Espion](01-points-arret-pas-a-pas-espion.md)** — Suspendre l'exécution, l'avancer instruction par instruction et surveiller la valeur des variables.
- **[19.2. Debug.Print, Assert et fenêtre Exécution immédiate](02-debug-print-assert.md)** — Tracer et inspecter par code, tester des expressions à la volée et poser des assertions de cohérence.
- **[19.3. Techniques de test des modules de données (DAO/ADO)](03-tests-modules-donnees.md)** — Vérifier le comportement du code d'accès aux données sans dépendre d'un état imprévisible.
- **[19.4. Tests unitaires en VBA — framework léger maison](04-tests-unitaires-framework.md)** — Construire un dispositif simple pour automatiser et rejouer des vérifications.
- **[19.5. Journalisation d'événements applicatifs](05-journalisation-evenements.md)** — Conserver une trace exploitable de l'activité pour diagnostiquer après coup, y compris en production.
- **[19.6. Simulation de données et environnements de test](06-simulation-donnees-test.md)** — Mettre en place des jeux de données et des environnements isolés, distincts de la production.
- **[19.7. Débogage à distance et techniques sur poste utilisateur](07-debogage-distance.md)** — Diagnostiquer un problème qui ne se reproduit que chez l'utilisateur, sans accès à l'éditeur.

## Prérequis

Ce chapitre suppose une bonne maîtrise des fondamentaux de VBA ([chapitre 3](../03-rappels-fondamentaux/README.md)) et de l'environnement de développement ([chapitre 2](../02-interface-environnement/README.md), notamment la [fenêtre Exécution immédiate](../02-interface-environnement/04-fenetres-proprietes-execution.md)). La gestion des erreurs ([chapitre 13](../13-gestion-erreurs/README.md)) en est le complément naturel, et les modules de classe ([chapitre 16](../16-poo-vba-access/README.md)) servent de base au framework de test. On notera que `Debug.Print`, abordé ici, est aussi l'outil de mesure de la [section 18.1](../18-optimisation-performance/01-profilage-mesure-performances.md).

## En résumé

Déboguer et tester, c'est refuser de programmer à l'aveugle : rendre le comportement observable, reproduire avant de corriger, et vérifier par des éléments concrets plutôt que par l'intuition. En développement, l'éditeur VBA offre pour cela des outils interactifs ; en production, la journalisation et les techniques à distance prennent le relais. Et parce que VBA ne fournit pas de framework, c'est au développeur de se donner les moyens d'automatiser ses vérifications et d'isoler ses tests des données réelles. Les sections suivantes déroulent cet outillage, du plus immédiat au plus structuré.

---


⏭️ [19.1. Points d'arrêt, exécution pas à pas et fenêtre Espion](/19-debogage-tests/01-points-arret-pas-a-pas-espion.md)
