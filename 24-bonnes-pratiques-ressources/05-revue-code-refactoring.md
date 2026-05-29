🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 24.5. Revue de code et refactoring d'applications existantes

Maintenant que le code peut être exporté en texte et placé sous gestion de versions (section 24.4), deux pratiques jusqu'ici difficiles dans Access deviennent enfin possibles : la **revue de code** et le **refactoring**. Toutes deux servent le même objectif — élever la qualité d'une application — mais sous deux angles complémentaires : relire pour prévenir, et restructurer pour assainir. Le titre précise « d'applications existantes », et ce n'est pas anodin : la réalité du développeur Access est rarement la page blanche. C'est bien plus souvent un héritage — une application tordue, bâtie au fil des ans par d'autres, dont dépend tout un service. Cette section explique comment relire et restructurer un tel code **sans le déstabiliser**.

## La revue de code : faire examiner avant d'intégrer

La revue de code consiste à faire examiner du code — par un pair ou par soi-même — afin de détecter les défauts, de partager la connaissance, et de garantir la conformité aux conventions (section 24.1) et à l'architecture (section 24.2). Ses bénéfices sont multiples : on attrape les bugs au plus tôt, on diffuse la compréhension de l'application au sein de l'équipe — désamorçant le « piège de l'unique sachant » évoqué en introduction — et on maintient une cohérence d'ensemble. Rare dans le monde Access, où elle alimente précisément le préjugé selon lequel les développeurs Access ne seraient pas de « vrais » développeurs, la revue de code est justement l'une des pratiques qui professionnalisent le plus le métier.

## Rendre la revue possible dans Access

La revue se heurtait au même obstacle que le versioning : on ne relit pas un fichier binaire. C'est l'**export texte et la mise sous Git** de la section précédente qui rendent la revue réalisable. On relit les différences (*diffs*) d'un commit, ou une demande de fusion (*pull request*) portant sur les fichiers source exportés. L'option consistant à séparer, pour chaque formulaire, la mise en page et le code VBA dans deux fichiers (section 24.4) prend ici tout son sens : elle permet de concentrer la revue sur les changements de logique, sans être noyé sous le texte de disposition. La revue de code est donc un prolongement direct de la gestion de versions.

## Ce qu'on regarde dans une revue

Une revue efficace ne s'éparpille pas ; elle balaie quelques axes essentiels. La **correction** d'abord : la logique est-elle juste, les cas limites traités, et surtout chaque procédure gère-t-elle ses erreurs (chapitre 13) ? Le `Option Explicit` est-il présent en tête de module ? Viennent ensuite les **conventions** : nommage (section 24.1), mise en forme, cohérence. Puis l'**architecture** : les formulaires restent-ils légers, sans règle métier ni SQL enfouis dans leurs événements (section 24.2), l'accès aux données est-il isolé ?

L'**accès aux données** mérite une attention particulière : les Recordsets sont-ils refermés et libérés (section 9.11), le SQL est-il paramétré pour prévenir les injections (section 11.5), les requêtes sont-elles efficaces (pas de `SELECT *` sur de grosses tables, filtrage côté serveur — chapitre 18) ? On vérifie enfin la **robustesse** (validation des saisies, section 8.10 ; transactions là où elles s'imposent, chapitre 14), la **lisibilité** (commentaires sur le pourquoi, section 24.3 ; absence de code mort et de nombres magiques) et la **sécurité** (aucun mot de passe en dur).

Sur la forme, une bonne revue suit quelques principes : on relit **le code, pas la personne** ; on reste constructif ; on privilégie les remarques à fort impact plutôt que le pinaillage ; et l'on retient que l'auto-revue — relire posément son propre code avant de le valider — est précieuse, même seul.

## Le refactoring : améliorer sans changer le comportement

Le refactoring consiste à **améliorer la structure interne du code sans modifier son comportement externe**. Les mêmes entrées produisent les mêmes sorties ; seul l'intérieur change : plus clair, mieux structuré, moins dupliqué. Cette définition impose une distinction stricte, trop souvent oubliée : refactoriser n'est **pas** ajouter une fonctionnalité, **pas** corriger un bug, **pas** réécrire. C'est une amélioration **à comportement constant**. Confondre ces gestes est la première cause de refactorings ratés.

## Le défi des applications existantes

Le code hérité présente un visage caractéristique : toute la logique dans les événements des formulaires, du code dupliqué, des noms cryptiques, du SQL disséminé, aucune gestion d'erreurs, aucun test. Le danger est double : le moindre changement peut tout casser, et il n'existe **aucun filet de sécurité** pour s'en apercevoir. D'où une règle d'ordre impérative : **établir la sécurité avant de refactoriser**. On ne se lance pas dans la restructuration d'un code hérité à mains nues.

## Refactoriser sans déstabiliser : la méthode

La restructuration sûre d'une application existante suit une démarche précise.

1. **Mettre le code sous gestion de versions d'abord** (section 24.4). C'est le prérequis numéro un : pouvoir observer chaque changement et revenir en arrière. On ne refactorise jamais sans contrôle de source.
2. **Établir un filet de sécurité.** Avant de toucher au code, on caractérise son comportement actuel — idéalement par des tests (sections 19.3 et 19.4), à défaut par des scénarios de test documentés ou une recette (section 24.6). Sur du code hérité sans tests, on écrit des *tests de caractérisation* qui capturent ce que le code fait *aujourd'hui* (juste ou non), afin de détecter toute régression ultérieure.
3. **Procéder par petits pas incrémentaux.** Une modification à la fois, vérifiée puis validée. Jamais de refonte massive en une seule fois : ce sont les changements réversibles et de petite taille qui maîtrisent le risque.
4. **Ne pas mélanger refactoring et changement de comportement.** On refactorise *ou* on ajoute une fonctionnalité, jamais les deux dans le même commit, afin qu'une régression soit toujours attribuable.
5. **Vérifier le comportement après chaque étape.** Le filet de sécurité n'a de valeur que si on l'utilise systématiquement.

## Les refactorings courants dans Access

Certains gestes reviennent constamment sur du code Access hérité. Le plus rentable est d'**extraire la logique métier des événements de formulaire** vers des modules et des classes, mouvement qui rapproche l'application de l'architecture en couches de la section 24.2.

```vba
' AVANT : règle métier et accès aux données dans l'événement
Private Sub cmdValider_Click()
    If Me.txtMontant > 1000 Then
        Me.txtRemise = Me.txtMontant * 0.05
    End If
    CurrentDb.Execute "INSERT INTO Journal (...) VALUES (...)"
End Sub

' APRÈS : le formulaire délègue ; la logique est extraite,
'         réutilisable et testable
Private Sub cmdValider_Click()
    Me.txtRemise = CalculerRemise(Me.txtMontant)
    EnregistrerValidation Me
End Sub
```

Dans le même esprit, on **centralise l'accès aux données** dans une couche dédiée (patron Repository, section 16.6) en remplaçant le SQL dispersé par des appels paramétrés (section 11.5) ; on **élimine les duplications** en extrayant le code répété dans des procédures réutilisables ; on **renomme** pour la clarté en appliquant les conventions de la section 24.1 ; on **ajoute `Option Explicit`** et l'on corrige les variables implicites que cela révèle ; on **ajoute une gestion d'erreurs structurée** (chapitre 13) là où elle manque ; on remplace les **nombres magiques** par des constantes nommées ; on substitue des **opérations ensemblistes** aux boucles enregistrement par enregistrement (chapitres 18 et 23.3) ; on **scinde** les modules géants et les formulaires surchargés ; et l'on **supprime le code mort**.

## Les outils

Plusieurs outils facilitent ce travail. **Rubberduck**, add-in open source pour VBA, apporte des inspections de code, des refactorings outillés (renommage, extraction de méthode) et des tests unitaires — un soutien précieux et spécifiquement pensé pour VBA. Les analyseurs comme **MZ-Tools** ou **Total Access Analyzer** repèrent le code mort, les variables non déclarées, les incohérences et la complexité, ce qui aide à **dresser l'état des lieux** d'une application héritée et à orienter le refactoring. À cela s'ajoute l'hygiène de décompilation et de compactage (section 24.4), utile pour partir d'une base saine. Les fonctionnalités exactes de ces outils évoluant, on se reportera à leur documentation à jour.

## La juste mesure : quand refactoriser, et refactoring contre réécriture

Le refactoring n'est pas une fin en soi. On ne restructure pas du code par perfectionnisme, mais lorsqu'on doit le modifier ou l'étendre, ou lorsque son désordre **bloque activement** le travail. Un code qui fonctionne et que l'on ne touche jamais peut très bien rester en l'état. Une approche équilibrée pour le code hérité est la « règle du scout » : laisser le code un peu plus propre qu'on ne l'a trouvé, au fil des interventions, plutôt que de tout assainir d'un coup.

Se pose enfin la tentation de la **réécriture complète**. L'expérience du développement logiciel est sans appel : les réécritures intégrales sont risquées et échouent souvent, car on sous-estime tout ce que le code existant a accumulé de corrections et de cas particuliers au fil des ans. Sauf situation extrême, il est presque toujours préférable de **refactoriser progressivement une application qui marche** que de la réécrire de zéro. La modernisation par petits pas est moins spectaculaire, mais infiniment plus sûre.

## En résumé

La revue de code et le refactoring, longtemps malaisés dans Access, deviennent praticables une fois le code exporté en texte et versionné (section 24.4). La revue fait examiner le code avant intégration, en se concentrant sur la correction, les conventions, l'architecture, l'accès aux données et la sécurité, dans un esprit constructif. Le refactoring améliore la structure interne sans changer le comportement — à condition, sur une application existante, de respecter un ordre strict : versionner d'abord, établir un filet de sécurité, procéder par petits pas, ne jamais mêler restructuration et changement de comportement, et vérifier à chaque étape. Les refactorings Access les plus rentables extraient la logique des formulaires, isolent l'accès aux données et éliminent les duplications, soutenus par des outils dédiés. La règle d'or reste la prudence : restructurer plutôt que réécrire, et toujours sous filet. Ce filet — la vérification du bon comportement avant livraison — est lui-même une discipline à part entière : c'est la recette, objet de la section suivante.

⏭️ [24.6. Checklist de recette avant livraison d'une application](/24-bonnes-pratiques-ressources/06-checklist-recette.md)
