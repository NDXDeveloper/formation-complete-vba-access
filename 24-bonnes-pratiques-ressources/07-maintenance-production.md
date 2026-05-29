🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 24.7. Maintenance et évolution d'une base Access en production

La livraison n'est pas la fin d'un projet : c'en est, le plus souvent, le début de la phase la plus longue. L'essentiel de la vie — et du coût — d'une application se joue après sa mise en service, pendant des années d'exploitation. Cela est particulièrement vrai pour Access, dont les applications, on l'a rappelé tout au long de ce chapitre, vivent longtemps. Maintenir, ce n'est d'ailleurs pas seulement corriger des bugs : c'est entretenir, adapter, améliorer et faire évoluer une application **vivante**, dont des utilisateurs dépendent quotidiennement. Et là réside la difficulté propre à la maintenance en production : il faut faire tout cela **sans interrompre** ceux qui s'en servent.

## Les quatre visages de la maintenance

La maintenance recouvre quatre réalités distinctes, qu'il est utile de nommer pour ne pas réduire le sujet aux seules corrections. La maintenance **corrective** répare les anomalies découvertes en exploitation. La maintenance **adaptative** ajuste l'application aux évolutions de son environnement : nouvelle version d'Office ou de Windows, changement de serveur, passage 32/64 bits. La maintenance **perfective** améliore l'existant sans en changer le comportement — optimisation (chapitre 18), refactoring (section 24.5). La maintenance **évolutive**, enfin, ajoute de nouvelles fonctionnalités. Une application en production sollicite ces quatre formes au fil de sa vie ; les confondre, ou n'en voir qu'une, conduit à mal organiser le travail.

## Ce qui rend la maintenance possible : les fondations posées

La maintenance sereine d'une application Access n'est possible que si les pratiques des sections précédentes sont en place. La **séparation front-end / back-end** (section 24.2) permet de mettre à jour l'application sans toucher aux données et sans interruption. La **gestion de versions** (section 24.4) trace chaque changement et autorise le retour en arrière. La **documentation** (section 24.3) permet de comprendre l'application qu'on maintient — souvent écrite par d'autres. La **gestion centralisée des erreurs et la journalisation** (sections 13.7 et 19.5) révèlent ce qui dysfonctionne réellement en production. Et la **recette** (section 24.6) valide chaque modification avant son redéploiement. Sans ces fondations, la maintenance se résume à du bricolage anxieux ; avec elles, elle devient une activité maîtrisée.

## La règle d'or : la production n'est pas un établi

S'il ne fallait retenir qu'une règle de la maintenance en production, ce serait celle-ci : **on ne modifie jamais directement l'application en production**. Toute modification, si minime soit-elle, suit le même cycle : on travaille sur une **copie de développement**, on la **versionne**, on la **teste**, on la passe en **recette**, puis on la **déploie**. La production n'est pas un établi sur lequel on bidouille à chaud, sous la pression d'un utilisateur qui attend. Idéalement, on dispose d'un environnement de **test ou de pré-production** reproduisant les conditions réelles (section 19.6), afin d'éprouver les changements avant qu'ils n'atteignent les utilisateurs. Enfreindre cette règle, c'est s'exposer tôt ou tard à corrompre des données ou à interrompre un service dont dépend tout un service.

## L'entretien courant

Au-delà des modifications, une application en production réclame un entretien régulier. La pratique la plus importante de toutes est la **sauvegarde** du back-end : fréquente, automatisée, et surtout **testée** — une sauvegarde qu'on n'a jamais su restaurer n'est pas une sauvegarde. On définit une politique de rétention et, idéalement, une copie sur un support distinct. Le **compactage et la réparation** périodiques du back-end (section 18.7) contiennent le gonflement du fichier et préservent sa santé, à mener prudemment, hors présence des utilisateurs ou pendant les heures creuses. À chaque redéploiement, on décompile et compacte aussi le front-end (section 24.4).

La **surveillance** complète cet entretien : on garde un œil sur la taille du fichier face à la limite des 2 Go (section 18.9), sur les performances, sur les journaux d'erreurs (section 19.5), et sur les fichiers de verrouillage orphelins (section 15.8). On reste enfin vigilant sur le **risque de corruption** — stabilité du réseau, fermetures propres — en gardant un plan de reprise sous la main. Lorsque le back-end est migré vers SQL Server (chapitre 23), une bonne part de cet entretien — sauvegardes, santé du moteur — devient la responsabilité du serveur, gérée avec ses propres outils.

## Déployer une mise à jour sans interrompre les utilisateurs

C'est le cœur de la maintenance en production. Parce que chaque utilisateur possède sa propre copie du front-end, **mettre à jour l'application ne touche pas aux données** : on déploie un nouveau front-end (via un mécanisme de mise à jour automatique, section 21.3), et chacun le reçoit à son prochain lancement, sans interruption du service de données. On programme ces déploiements en heures creuses et on communique aux utilisateurs.

La difficulté se concentre sur les **changements de structure du back-end**. Ajouter un champ ou une table affecte tout le monde **simultanément** et exige une coordination : on opère dans une fenêtre de maintenance, après sauvegarde, on teste la reliaison, et l'on s'assure que le nouveau front-end correspond bien au nouveau schéma. Un excellent garde-fou consiste à **versionner le schéma** et à faire **vérifier la compatibilité au démarrage** par le front-end, qui refuse de s'exécuter en cas de décalage — évitant ainsi qu'un front-end désynchronisé ne provoque erreurs ou corruptions sur un schéma qui ne lui correspond pas.

```vba
' Au démarrage : vérifier que ce front-end est compatible
' avec la version du schéma du back-end avant d'ouvrir l'application
Private Const VERSION_FRONT_END As Long = 12

Public Function VerifierCompatibilite() As Boolean
    Dim lngVersionBackEnd As Long
    lngVersionBackEnd = Nz(DLookup("VersionSchema", "tblParametres"), 0)

    If lngVersionBackEnd <> VERSION_FRONT_END Then
        MsgBox "Cette version de l'application n'est pas compatible " & _
               "avec la base de données. Veuillez installer la mise à jour.", _
               vbCritical, "Mise à jour requise"
        VerifierCompatibilite = False
    Else
        VerifierCompatibilite = True
    End If
End Function
```

## Faire évoluer l'application

L'évolution suit les mêmes règles que toute modification, sans exception. Ajouter une fonctionnalité passe par le cycle discipliné : une branche (section 24.4), du développement, des tests (chapitre 19), une recette (section 24.6), un déploiement (section 21.3) — jamais une retouche improvisée sur le fichier de production.

La maintenance adaptative demande, elle, une vigilance dans la durée. Une nouvelle version d'Office ou de Windows, un changement de bitness (section 21.7), une référence devenue « manquante » (section 2.5), une déclaration d'API à mettre à jour (section 22.1) : il faut périodiquement vérifier que l'application tourne toujours sur les environnements courants, et anticiper les transitions. Lorsque l'application ne cesse de croître — en volume, en utilisateurs, en complexité — la bonne réponse n'est pas toujours d'ajouter encore : c'est parfois de reconsidérer son architecture (section 24.2), ses performances (chapitre 18), voire d'envisager une migration (chapitre 23).

## Préserver la continuité de la connaissance

Une application en production survit généralement à la présence de celui qui l'a écrite. La maintenance durable suppose donc de **préserver la connaissance** : la documentation (section 24.3) et la gestion de versions (section 24.4) font que l'entretien ne repose pas sur une seule tête — c'est l'antidote au « piège de l'unique sachant » qui traverse tout ce chapitre. Un **processus de support** structuré complète le dispositif : un numéro de version visible dans l'application (section 24.6) pour identifier ce qui tourne réellement, une voie claire pour que les utilisateurs signalent les anomalies, et des journaux (section 19.5) pour reproduire et diagnostiquer les incidents.

## Quand la maintenance annonce une migration

Certains signaux relevés en exploitation dépassent la maintenance ordinaire : corruptions répétées, fichier qui approche des 2 Go, lenteurs et blocages liés à la concurrence. Ces observations ne sont pas que des problèmes à traiter ; ce sont les **indicateurs** qui alimentent la décision de migrer (section 23.1). Savoir reconnaître qu'une application a dépassé le cadre d'Access fait partie intégrante de sa maintenance. La boucle rejoint ainsi le chapitre consacré à la migration : maintenir, c'est aussi savoir quand il est temps de changer d'échelle.

## En résumé

La maintenance est la phase la plus longue de la vie d'une application, et la plus déterminante pour sa pérennité. Elle revêt quatre visages — corrective, adaptative, perfective, évolutive — et ne devient sereine que si les fondations sont posées : séparation front-end/back-end, gestion de versions, documentation, journalisation, recette. Sa règle d'or est que la production n'est pas un établi : toute modification passe par une copie, des tests et une recette. L'entretien courant repose avant tout sur des sauvegardes testées et un compactage régulier ; les mises à jour se déploient par le front-end, sans interruption, tandis que les changements de schéma se coordonnent en fenêtre de maintenance, protégés par une vérification de version. L'évolution suit le cycle discipliné, la connaissance se préserve par la documentation et le versioning, et les signaux de saturation annoncent, le cas échéant, une migration. Bien mener tout cela est plus aisé quand on n'est pas seul : c'est pourquoi la dernière section du chapitre, et de la formation, est consacrée à la communauté et aux ressources.

⏭️ [24.8. Communauté et ressources en ligne](/24-bonnes-pratiques-ressources/08-communaute-ressources.md)
