🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 24.6. Checklist de recette avant livraison d'une application

La **recette** est la phase de validation qui précède la mise en service : on y vérifie méthodiquement que l'application est conforme à ce qui était attendu et prête à être livrée. Elle peut être un acte formel — un procès-verbal de recette par lequel le client valide la conformité au cahier des charges — ou une simple discipline interne avant la bascule en production. Dans tous les cas, son principe est le même : on ne livre pas sur la foi d'un « ça avait l'air de marcher ». La recette est le prolongement naturel du filet de sécurité évoqué à la section précédente, appliqué à l'échelle de toute l'application. Cette section propose une checklist structurée, à adapter à chaque projet.

## Trois principes avant la liste

Avant le détail, trois principes pèsent plus lourd que tous les points de contrôle réunis.

**Tester sur une machine propre, représentative de la cible.** C'est la règle d'or, et la première cause d'échec des déploiements Access est de l'ignorer. Une application qui fonctionne sur le poste de développement — où tout est installé, configuré, approuvé — peut échouer lamentablement ailleurs. La recette doit se dérouler sur un environnement vierge, conforme à celui des utilisateurs (version d'Access ou Runtime, bitness, droits).

**Tester avec un volume de données réaliste.** Les quelques enregistrements de développement ne révèlent ni les problèmes de performance, ni certains comportements limites. On valide avec un jeu de données proche du réel (section 19.6).

**Valider contre le comportement attendu, pas seulement contre l'absence de plantage.** Recette signifie conformité : l'application fait-elle *ce qu'elle doit faire*, conformément au cahier des charges ? « Ça ne plante pas » n'est pas « c'est correct ».

La liste qui suit est un **gabarit** : tous les points ne s'appliquent pas à toutes les applications, et il convient de l'ajuster au contexte.

## Code et compilation

- [ ] L'application compile sans erreur (Déboguer → Compiler)
- [ ] `Option Explicit` figure en tête de chaque module
- [ ] Aucun code de débogage résiduel (`Debug.Print`, `Stop`, `MsgBox` de test, points d'arrêt)
- [ ] Décompilation puis compactage/réparation effectués (section 24.4)
- [ ] Gestion d'erreurs présente dans chaque procédure, avec un gestionnaire global (chapitre 13)
- [ ] Aucun chemin, nom de serveur ou mot de passe codé en dur ; configuration centralisée (section 24.2)

## Données et intégrité

- [ ] Séparation front-end / back-end effective ; aucune table locale parasite (section 24.2)
- [ ] Reliaison des tables au démarrage fonctionnelle (sections 21.4 et 23.4), testée sur poste vierge
- [ ] Relations et intégrité référentielle définies (section 14.6) ; index et champs requis en place
- [ ] Aucune donnée de test résiduelle ; données de référence présentes et correctes
- [ ] Comportement des compteurs (NuméroAuto / IDENTITY) vérifié, surtout avec un back-end SQL Server (section 23.3)
- [ ] Sauvegarde du back-end réalisée avant la bascule ; stratégie de sauvegarde définie (section 24.7)

## Fonctionnel

- [ ] Chaque formulaire principal s'ouvre et permet création, lecture, modification et suppression
- [ ] Chaque état s'imprime, y compris en l'absence de données (gestion NoData, section 7.9)
- [ ] Navigation complète : tous les menus et boutons mènent à un état valide
- [ ] Règles métier conformes au cahier des charges (section 24.2)
- [ ] Recherches, filtres, calculs et totaux corrects
- [ ] Cas limites testés : saisies vides, valeurs extrêmes, caractères spéciaux, bornes de dates
- [ ] Scénarios multi-utilisateurs testés si applicable : accès concurrents, verrouillage (chapitre 15)

## Interface et ergonomie

- [ ] Ordre de tabulation correct ; navigation au clavier fonctionnelle (section 17.6)
- [ ] Libellés exacts, sans fautes, langue cohérente
- [ ] Messages d'erreur clairs et compréhensibles par l'utilisateur (section 8.10)
- [ ] Affichage correct à la résolution cible ; aucun contrôle tronqué ; apparence cohérente (section 17.9)
- [ ] Options de démarrage réglées : titre, icône, volet de navigation, touche d'évitement (section 17.7)

## Performance

- [ ] Ouverture des formulaires et états dans un délai acceptable
- [ ] Formulaires sur grosses tables filtrés (section 18.5) ; aucune requête `SELECT *` inutile (chapitre 18)
- [ ] Performances validées avec un volume réaliste (section 19.6) et contre le vrai back-end (réseau)

## Sécurité et protection

- [ ] Diffusion en `.accde` si le code doit être verrouillé (section 20.3)
- [ ] Aucun secret en dur ; données sensibles protégées (section 20.8)
- [ ] Droits utilisateurs appliqués si applicable (section 20.6)
- [ ] L'application s'exécute sans avertissement de sécurité dans l'environnement cible (section 20.5) ; signée si requis (section 20.4)

## Déploiement et environnement

- [ ] Testé sur une machine propre représentative de la cible — pas seulement le poste de développement
- [ ] Références présentes, aucune marquée « MANQUANTE » (section 2.5)
- [ ] Pilotes ODBC présents et de bon bitness si back-end SQL Server ou autre (sections 23.4 et 23.7)
- [ ] Mécanisme de mise à jour et de déploiement du front-end fonctionnel (section 21.3)
- [ ] Compatibilité Runtime vérifiée si déploiement sans licence complète (section 21.2)
- [ ] Compatibilité 32/64 bits des déclarations API (sections 22.1 et 21.7) si des API Windows sont utilisées

## Documentation et passation

- [ ] Code documenté (section 24.3) ; documentation de projet (README) à jour
- [ ] Numéro de version défini et visible dans l'application (pour le support)
- [ ] Notes de version et journal des changements rédigés
- [ ] Documentation utilisateur fournie si nécessaire
- [ ] Source validée et étiquetée (tag) dans Git (section 24.4)

## Robustesse et derniers contrôles

- [ ] Aucune boîte d'erreur VBA brute exposée à l'utilisateur (gestionnaire global, section 13.7)
- [ ] Journalisation opérationnelle pour le diagnostic post-livraison (section 19.5)
- [ ] Fermeture propre : Recordsets fermés, connexions libérées
- [ ] Routine de démarrage (AutoExec) validée de bout en bout au premier lancement

## Adapter la checklist

Cette liste n'a pas vocation à être suivie aveuglément. Une petite application mono-utilisateur ignorera les points multi-utilisateurs et une bonne part des contrôles de déploiement ; une application critique en réseau, à l'inverse, justifiera d'en ajouter (procédures de retour arrière, fenêtre de bascule, communication aux utilisateurs). L'essentiel est d'en faire **sa** checklist, de la **versionner avec le projet** (section 24.4), et de l'**enrichir** au fil des livraisons : chaque incident en production qui aurait pu être attrapé à la recette mérite d'y devenir un nouveau point de contrôle. Une checklist de recette est un document vivant, qui s'affine avec l'expérience.

## En résumé

La recette est la vérification systématique d'une application avant sa livraison — formelle ou interne, elle interdit de livrer à l'aveugle. Trois principes la gouvernent : tester sur une machine propre représentative de la cible, avec un volume de données réaliste, et valider la conformité au comportement attendu et non la seule absence de plantage. La checklist proposée couvre le code et la compilation, les données et l'intégrité, le fonctionnel, l'interface, la performance, la sécurité, le déploiement, la documentation et la robustesse — autant de points à cocher, à adapter à chaque projet et à enrichir à chaque incident évité. Elle constitue la porte d'entrée de la production. Car une fois l'application livrée, elle ne cesse pas de vivre : elle entre dans une phase d'exploitation faite de surveillance, de corrections et d'évolutions — la maintenance, objet de la section suivante.

⏭️ [24.7. Maintenance et évolution d'une base Access en production](/24-bonnes-pratiques-ressources/07-maintenance-production.md)
