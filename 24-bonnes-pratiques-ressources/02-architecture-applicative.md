🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 24.2. Architecture applicative recommandée pour Access

La section précédente traitait du nommage, c'est-à-dire de l'ordre au niveau du détail. Celle-ci s'élève d'un cran : il s'agit de l'ordre au niveau de l'ensemble, l'**architecture** de l'application. Access a ceci de particulier qu'il rend la désorganisation extrêmement facile : rien n'empêche d'écrire toute la logique dans les événements des formulaires, de disséminer des chaînes SQL aux quatre coins du projet, et d'entasser le code dans un unique module. Le résultat est l'application ingérable décrite dans l'introduction du chapitre. Une architecture réfléchie est précisément ce qui évite cette dérive — et ce qui rend possibles la maintenance, les tests, le travail à plusieurs, et la migration éventuelle vers un autre moteur.

## La première décision : séparer le front-end du back-end

Pour toute application Access non triviale, la première décision architecturale n'est pas négociable : **séparer les données du reste de l'application**. Les données vivent dans un fichier *back-end* qui ne contient que des tables ; l'application — formulaires, états, requêtes, modules — vit dans un fichier *front-end* qui accède aux tables par liaison. C'est le principe de la base scindée, déjà rencontré au chapitre 15.

Les bénéfices sont considérables et conditionnent tout le reste. Chaque utilisateur reçoit sa **propre copie du front-end**, ce qui rend le multi-utilisateur sain et évite de partager un fichier unique. Les mises à jour de l'application se déploient en remplaçant le front-end, sans jamais toucher aux données (chapitre 21). Le risque de corruption diminue. Et surtout, cette séparation rend l'application **prête à migrer** : remplacer le back-end Access par SQL Server (chapitre 23) ne change pas le principe. Cette première décision est donc le socle de toute architecture durable.

## Au-delà du fichier : séparer les responsabilités en couches

La séparation front-end/back-end est *physique*. Il en faut une seconde, *logique*, à l'intérieur même du front-end : organiser le code en **couches** aux responsabilités distinctes, comme exposé au chapitre 16. On distingue trois couches.

```
┌─────────────────────────────────────────────────────┐
│  Présentation (UI)                                  │
│  Formulaires, états, gestionnaires d'événements     │
│  -> « légers » : ils délèguent, ils ne décident pas │
└────────────────────┬────────────────────────────────┘
                     │ appelle
┌────────────────────▼────────────────────────────────┐
│  Logique métier (BLL)                               │
│  Règles, validations, calculs, workflows            │
│  Modules standard et classes du domaine             │
└────────────────────┬────────────────────────────────┘
                     │ sollicite
┌────────────────────▼────────────────────────────────┐
│  Accès aux données (DAL)                            │
│  Recordsets, SQL, Repository                        │
│  -> seul à connaître réellement la base             │
└────────────────────┬────────────────────────────────┘
                     │ DAO / ADO / ODBC
               ┌──────▼────────┐
               │  Back-end     │
               │  (Access ou   │
               │   SQL Server) │
               └───────────────┘
```

Le principe directeur tient en une phrase : **chaque couche ne parle qu'à sa voisine immédiate**, et chacune a une responsabilité unique. La présentation affiche et collecte, la logique métier décide, l'accès aux données lit et écrit. Tout l'enjeu est d'éviter que ces rôles ne se mélangent — ce qui est précisément la pente naturelle d'Access.

## La couche de présentation : des formulaires « légers »

Un formulaire devrait afficher des données et capter des actions, rien de plus. Les règles métier et les accès à la base n'ont pas leur place dans ses gestionnaires d'événements : ceux-ci doivent **déléguer** le travail aux couches inférieures.

```vba
' À éviter : règle métier et SQL noyés dans l'événement du formulaire
Private Sub cmdValider_Click()
    If Me.txtMontant > 10000 And Me.cboCategorie = "Standard" Then
        ' ... traitement ...
    End If
    CurrentDb.Execute "UPDATE Commandes SET Statut='V' WHERE ..."
End Sub

' Préférer : le formulaire délègue à la couche métier
Private Sub cmdValider_Click()
    If ValiderCommande(Me) Then
        EnregistrerCommande Me
    End If
End Sub
```

Un formulaire « léger » est plus lisible, et surtout la logique qu'il déclenche devient réutilisable et testable, car elle ne dépend plus de l'écran. On veillera par ailleurs à ne pas lier aveuglément un formulaire à une table entière, mais à filtrer sa source (section 18.5) — l'architecture rejoint ici la performance.

## La couche métier : les règles au cœur, indépendantes de l'écran

La couche métier rassemble ce qui fait la valeur de l'application : ses **règles de gestion, ses validations, ses calculs et ses workflows**. Ce code réside dans des modules standard et, mieux encore, dans des **classes du domaine** représentant les objets métier (un client, une commande, une facture). Placé là plutôt que dans les formulaires, il devient indépendant de l'interface : on peut le réutiliser depuis plusieurs écrans, l'appeler depuis un traitement par lots, et le tester isolément (chapitre 19). C'est dans cette couche que se concentre l'intelligence de l'application.

## La couche d'accès aux données : isoler le dialogue avec la base

La troisième couche encapsule tout ce qui touche réellement à la base : ouverture de Recordsets, construction de SQL, lecture et écriture. L'anti-pattern à fuir est la dispersion : du SQL et des appels DAO disséminés dans des dizaines de formulaires et de modules. Au contraire, on **centralise** l'accès aux données, idéalement via le patron Repository (section 16.6), de sorte que seul ce point connaisse les détails de la base.

Ce point d'isolement est ce qui rend une application **migrable**. Le jour où le back-end Access devient SQL Server (chapitre 23), c'est cette seule couche qu'il faut adapter — pas les centaines de formulaires. C'est aussi ce qui permet de tester la logique métier sans vraie base, et de garantir la cohérence des accès. La couche d'accès aux données est donc bien plus qu'un détail technique : c'est une assurance sur l'avenir de l'application.

## L'organisation du code en modules

À l'échelle du code, l'architecture se traduit par une organisation claire des modules, à l'opposé de l'unique « Module1 » fourre-tout. On répartit le code **par responsabilité** : un module de démarrage, un module d'utilitaires, un module de gestion centralisée des erreurs (section 13.7), des modules d'accès aux données, des modules métier ; et des **modules de classe** pour les entités du domaine et les services. Cette répartition rend chaque fonctionnalité localisable, et constitue la traduction concrète, au niveau des fichiers de code, des couches décrites plus haut.

## L'infrastructure transversale

Toute application sérieuse repose en outre sur une **ossature commune**, indépendante du métier, qu'il vaut mieux concevoir dès le départ. Une **routine de démarrage** initialise l'application : elle reconnecte les tables liées (sections 21.4 et 23.4), vérifie la version, prépare les variables globales et les `TempVars`, puis ouvre l'interface principale. Un **module de gestion centralisée des erreurs** (section 13.7) traite les erreurs de façon uniforme. Un mécanisme de **configuration** — table de paramètres, `TempVars` ou registre — centralise les réglages plutôt que de les coder en dur. Une **journalisation** (section 19.5) trace les événements applicatifs. Enfin, un **point de navigation** central (formulaire de navigation, section 17.3) organise l'accès aux écrans. Cette infrastructure, mise en place une fois, profite à toute l'application.

## Pourquoi l'architecture compte particulièrement dans Access

Il faut redire pourquoi ce soin est plus crucial dans Access qu'ailleurs. La plupart des environnements modernes encouragent, par leur structure même, une certaine séparation des responsabilités. Access fait l'inverse : la voie de moindre résistance y consiste à tout écrire dans les formulaires et à coller du SQL partout. C'est rapide, cela fonctionne… et cela produit, à terme, la « grosse boule de boue » que personne n'ose plus toucher.

Une architecture délibérée est l'antidote. C'est elle qui rend possible la **migration** vers un vrai moteur (chapitre 23), qui permet les **tests** (chapitre 19), qui autorise le **travail en équipe**, et qui désamorce le « piège de l'unique sachant » évoqué en introduction. Investir dans l'architecture, ce n'est pas un luxe d'esthète : c'est ce qui détermine si l'application sera maintenable dans cinq ans ou non.

## Ne pas sur-concevoir : adapter à la taille

Cette rigueur doit toutefois rester proportionnée. Une petite application de quelques formulaires n'a pas besoin de l'appareillage complet du patron Repository, des classes de domaine et de couches strictement étanches : ce serait de la sur-ingénierie, aussi nuisible que l'absence d'architecture. La bonne mesure dépend de la **taille et de la durée de vie** attendues de l'application.

Cela dit, deux principes restent valables quelle que soit l'échelle : **scinder front-end et back-end**, et **ne pas enfouir toute la logique dans les formulaires**. Ce sont les deux acquis minimaux. Au-delà, on enrichit l'architecture à proportion des besoins. L'objectif n'est jamais la pureté architecturale pour elle-même, mais la **maintenabilité** — la capacité, pour soi et pour les autres, de comprendre et de faire évoluer l'application sans crainte.

## En résumé

L'architecture recommandée pour Access repose sur deux séparations : une séparation physique entre le front-end (l'application, un exemplaire par utilisateur) et le back-end (les données, Access ou SQL Server) ; et une séparation logique en couches — présentation, logique métier, accès aux données — où chaque couche a une responsabilité unique. Les formulaires restent légers et délèguent, la logique métier se concentre dans des modules et des classes indépendants de l'écran, et l'accès aux données est isolé, ce qui rend l'application testable et migrable. Une infrastructure transversale (démarrage, erreurs, configuration, journalisation, navigation) complète l'ensemble. Le tout doit être dosé selon la taille de l'application, sans dogmatisme, avec pour seule boussole la maintenabilité. Une fois l'application bien structurée, encore faut-il qu'elle soit compréhensible : c'est le rôle de la documentation, objet de la section suivante.

⏭️ [24.3. Documentation du code et génération de documentation](/24-bonnes-pratiques-ressources/03-documentation-code.md)
