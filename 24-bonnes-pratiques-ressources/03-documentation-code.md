🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 24.3. Documentation du code et génération de documentation

Documenter, c'est communiquer à travers le temps. Le destinataire d'un commentaire n'est presque jamais soi-même au moment où on l'écrit : c'est un collègue qui reprendra l'application, un successeur des années plus tard, ou soi-même ayant tout oublié. Comme le rappelait l'introduction du chapitre, les applications Access vivent longtemps et changent de mains ; la documentation est ce qui empêche qu'elles ne deviennent des énigmes. Cette section traite des deux faces de la question : bien documenter le code, et générer de la documentation à partir de l'application. Elle s'appuie directement sur les deux sections précédentes — un code bien nommé (24.1) et bien structuré (24.2) est déjà, en lui-même, à demi documenté.

## Le meilleur commentaire est un code qu'on n'a pas besoin de commenter

Avant de parler de commentaires, il faut rappeler une évidence souvent oubliée : la première documentation est le code lui-même. Des noms explicites et une architecture claire rendent inutile une bonne partie des explications. Une fonction nommée `CalculerRemiseClient` avec des paramètres clairs se passe de commentaire pour dire ce qu'elle fait. Les commentaires ne **remplacent** donc pas un bon code ; ils le **complètent** là où il ne peut parler de lui-même. C'est cet état d'esprit qui doit guider tout le reste.

## Commenter le pourquoi, pas le quoi

De là découle la règle cardinale du commentaire : expliquer le **pourquoi**, pas le **quoi**. Le code dit déjà ce qu'il fait ; un commentaire qui le répète n'ajoute que du bruit — et un bruit dangereux, car il finit par mentir lorsque le code évolue sans que le commentaire soit mis à jour.

```vba
' Inutile : le commentaire répète le code
i = i + 1   ' incrémente i

' Utile : le commentaire explique le pourquoi
' On force le recalcul ici car Access ne rafraîchit pas
' le sous-formulaire automatiquement après un Requery du parent.
Me.sfDetails.Requery
```

Les commentaires précieux sont ceux qui éclairent ce que le code ne peut dire : l'**intention**, les **règles métier** et leur justification, les **décisions non évidentes**, les **contournements** d'un bug d'Access, les **hypothèses** et les unités. Un corollaire impératif en découle : un commentaire faux est pire que pas de commentaire. La documentation doit être **maintenue** au même titre que le code. On marquera enfin de façon cohérente les points à traiter (`TODO`, `FIXME`) pour les retrouver aisément.

## Les en-têtes de procédures et de modules

VBA ne dispose pas d'un système de documentation structurée intégré, comme les commentaires XML d'autres langages. On s'appuie donc sur une **convention d'en-tête** : un bloc placé en tête de chaque procédure, décrivant son objet, ses paramètres, sa valeur de retour, ses éventuels effets de bord, et son historique de révision. Chaque module reçoit de même un en-tête résumant son rôle.

```vba
'------------------------------------------------------------
' Procédure : CalculerRemise
' Objet     : Calcule le taux de remise d'une commande selon
'             le montant et la catégorie du client.
' Paramètres:
'   dblMontant   - Montant HT de la commande
'   strCategorie - Catégorie ("Standard" ou "Premium")
' Retour    : Double - taux de remise (0 à 1)
' Remarques : Seuils validés par le service commercial.
'             Ne pas modifier sans validation.
' Auteur    : ...            Date : ...
'------------------------------------------------------------
Public Function CalculerRemise(ByVal dblMontant As Double, _
                               ByVal strCategorie As String) As Double
    ' ...
End Function
```

La valeur de ces en-têtes tient à leur **cohérence** : un même gabarit, appliqué partout. On l'insère commodément au moyen d'un snippet (annexe K) ou d'un outil dédié. Cette discipline va de pair avec le `Option Explicit` en tête de chaque module, qui rend les intentions explicites et prévient toute une classe d'erreurs.

## Le dictionnaire de données : la propriété Description

Une astuce propre à Access mérite d'être systématisée : renseigner la propriété **Description** des champs (et des objets). Ce texte n'est pas décoratif — il remonte dans la barre d'état lors de la saisie, dans le Documenteur, et dans les vues de structure. Renseigner la description de chaque champ revient ainsi à constituer un **dictionnaire de données vivant**, intégré au fichier lui-même, à un coût minime pour un bénéfice durable. C'est l'une des formes de documentation les plus rentables dans Access.

## Générer la documentation : le Documenteur intégré

Access embarque un outil de génération de documentation, le **Documenteur de base de données** (Outils de base de données → Analyser → Documenteur). Il produit un rapport détaillé des objets sélectionnés : tables et champs avec leurs propriétés, requêtes, formulaires, états, relations et autorisations. C'est un bon moyen d'obtenir un **instantané de la structure** de l'application, exportable et archivable. Ses limites sont à connaître : le rapport est très **verbeux**, statique, et peu adapté à la documentation du code proprement dit. Il excelle pour documenter le **modèle de données**, beaucoup moins pour la logique.

## Générer la documentation par code

Pour une documentation sur mesure, on peut exploiter le modèle objet d'Access. Parcourir les collections `TableDefs`, `QueryDefs`, `AllForms` permet de produire automatiquement, par exemple, un dictionnaire de données au format Markdown ou HTML, régénérable à volonté et toujours à jour. Pour documenter le **code** lui-même, le modèle d'extensibilité de l'éditeur VBA (référence *Microsoft Visual Basic for Applications Extensibility*, et option « Accès approuvé au modèle d'objet du projet VBA » activée — voir chapitre 2) donne accès aux modules et aux procédures, dont on peut extraire les signatures et les en-têtes par programmation. On dispose ainsi d'une documentation générée à partir de la source, qui ne se périme pas tant qu'on la régénère.

## Les outils tiers

Au-delà des moyens intégrés, des **add-ins de productivité** automatisent ce travail. Des outils établis comme MZ-Tools ou Total Access Analyzer insèrent des en-têtes standardisés (garantissant la cohérence), génèrent de la documentation, analysent le code, repèrent le code mort ou les incohérences. Pour une application conséquente ou une équipe, ils font gagner un temps réel et imposent une régularité difficile à tenir à la main. Les fonctionnalités exactes évoluant, on se référera à leur documentation à jour ; l'essentiel est de savoir que cette catégorie d'outils existe et qu'elle vaut l'investissement sur les gros projets.

## La documentation externe : le grand tableau

Tout ne se documente pas dans le code. Une application a besoin d'une **documentation de projet** — un fichier README, un wiki — décrivant la vue d'ensemble : à quoi sert l'application, quelle est son architecture (section 24.2), comment est organisé le modèle de données, comment se déroulent le déploiement et la reliaison des tables, quelles sont les dépendances et les anomalies connues. Cette documentation du « grand tableau » a vocation à vivre **à côté du code**, dans le système de gestion de versions (section 24.4). C'est elle qui permet à un nouvel intervenant de comprendre l'application avant même d'en ouvrir le code.

## La juste mesure

Comme pour l'architecture, l'excès nuit. Une documentation pléthorique, qui paraphrase chaque ligne, n'est pas lue, n'est pas maintenue, et se transforme en piège dès qu'elle diverge du code. La bonne approche est ciblée : d'abord un code qui se documente lui-même par ses noms et sa structure, puis des commentaires limités au pourquoi et au non-évident, des en-têtes cohérents, un dictionnaire de données via les descriptions, une documentation structurelle générée, et un README pour la vue d'ensemble. Le tout maintenu — une documentation périmée est plus nuisible qu'absente.

## En résumé

Documenter, c'est transmettre à travers le temps une application qui survivra à son auteur. La meilleure documentation reste un code bien nommé et bien structuré ; les commentaires le complètent en expliquant le pourquoi et le non-évident, jamais le quoi. Des en-têtes cohérents sur les procédures et les modules pallient l'absence de système intégré dans VBA, et la propriété Description des champs constitue un dictionnaire de données vivant propre à Access. Côté génération, le Documenteur intégré offre un instantané de la structure, que le code et les outils tiers permettent d'enrichir et d'automatiser, tandis qu'un README externe porte la vue d'ensemble. Le maître mot demeure la juste mesure, et la maintenance de ce qu'on documente. Cette documentation externe, comme le code lui-même, gagne à vivre sous gestion de versions — ce qui pose la question, propre à Access en raison de son format binaire, du versioning avec Git, objet de la section suivante.

⏭️ [24.4. Versioning d'une application Access avec Git (décompilation, export texte)](/24-bonnes-pratiques-ressources/04-versioning-git.md)
