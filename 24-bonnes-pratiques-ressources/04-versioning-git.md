🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 24.4. Versioning d'une application Access avec Git (décompilation, export texte)

Mettre une application sous gestion de versions est, dans le monde du développement, une évidence : on suit l'historique des changements, on revient en arrière en cas d'erreur, on collabore, on relit. Pourtant, l'adoption de cette pratique reste étonnamment faible dans la communauté Access, et pour une raison technique précise : **une application Access est un fichier binaire**, là où Git est conçu pour le texte. Cette section explique comment surmonter cet obstacle, grâce aux deux techniques annoncées par son titre — l'export texte et la décompilation — et comment bâtir un véritable workflow Git autour d'une application Access. C'est, après le nommage, l'architecture et la documentation, une étape de plus vers le traitement d'Access comme un vrai logiciel.

## Le problème : un fichier binaire dans un monde de texte

Un fichier `.accdb` est un conteneur binaire unique qui renferme, sérialisés ensemble, les tables, requêtes, formulaires, états, macros et modules. Git, lui, raisonne en lignes de texte : c'est ainsi qu'il calcule des différences (*diffs*) lisibles, permet de fusionner des modifications, retrace l'auteur de chaque ligne et construit un historique signifiant.

Que se passe-t-il si l'on se contente de placer le `.accdb` binaire dans Git ? On obtient des **sauvegardes** : Git stocke des copies successives du fichier, et l'on peut revenir à une version antérieure. Mais on n'obtient pas de **gestion de versions** au sens plein : aucune différence lisible (impossible de voir *ce qui* a changé dans le code d'un formulaire d'un commit à l'autre), aucune fusion possible (deux développeurs ne peuvent pas combiner leurs changements ; un conflit binaire est insoluble), un dépôt qui enfle à chaque commit, et aucun historique ligne à ligne. Le binaire dans Git, c'est du *backup*, pas du *versioning*.

## La solution : exporter les objets en texte

La clé consiste à **exporter les objets de l'application en fichiers texte**, et à versionner ces fichiers plutôt que le binaire. Une fois le code et les définitions d'objets sous forme de texte, Git retrouve toute sa puissance : différences lisibles, fusion, historique et *blame* par objet.

### SaveAsText et LoadFromText

Le socle de cette approche tient en deux méthodes de l'objet `Application` : **`SaveAsText`** exporte un objet vers un fichier texte, et **`LoadFromText`** reconstruit l'objet à partir de ce texte. Longtemps non documentées officiellement — bien que toujours présentes et désormais reconnues par la documentation Microsoft — elles sont le fondement universel du contrôle de source dans Access. Elles s'appliquent aux formulaires, états, requêtes, macros et modules.

```vba
' Exporter les objets de l'application vers un dossier texte versionnable
Public Sub ExporterVersTexte(ByVal strDossier As String)
    Dim obj As AccessObject

    ' Modules (code VBA : modules standard et de classe)
    For Each obj In CurrentProject.AllModules
        Application.SaveAsText acModule, obj.Name, _
            strDossier & "\" & obj.Name & ".bas"
    Next obj

    ' Formulaires
    For Each obj In CurrentProject.AllForms
        Application.SaveAsText acForm, obj.Name, _
            strDossier & "\Form_" & obj.Name & ".txt"
    Next obj

    ' États
    For Each obj In CurrentProject.AllReports
        Application.SaveAsText acReport, obj.Name, _
            strDossier & "\Report_" & obj.Name & ".txt"
    Next obj

    ' Requêtes
    For Each obj In CurrentData.AllQueries
        Application.SaveAsText acQuery, obj.Name, _
            strDossier & "\Query_" & obj.Name & ".txt"
    Next obj

    ' Macros
    For Each obj In CurrentProject.AllMacros
        Application.SaveAsText acMacro, obj.Name, _
            strDossier & "\Macro_" & obj.Name & ".txt"
    Next obj
End Sub
```

À noter que les **tables et leurs données** ne se versionnent pas par ce biais : `SaveAsText` ne couvre pas les données. On versionne le *schéma* (la structure), pas le contenu — d'autant que les données vivent dans le back-end et n'ont pas vocation à figurer dans Git. Les outils spécialisés gèrent les définitions de tables par des moyens dédiés.

## La décompilation : repartir d'une source propre

Le titre associe à juste titre le versioning et la **décompilation**, car tous deux touchent à la dualité entre le code source et sa forme compilée. Access conserve en effet le code VBA sous deux formes : le **texte source** et un **code compilé** (p-code). Avec le temps, cet état compilé peut se corrompre, gonfler et devenir incohérent, provoquant des comportements erratiques, des erreurs de compilation mystérieuses et une instabilité.

La commande `/decompile`, passée en ligne de commande, **élimine le code compilé** et force Access à le régénérer à partir de la source.

```
"C:\Program Files\...\Office16\MSACCESS.EXE" "C:\App\MonAppli.accdb" /decompile
```

(le chemin exact de `MSACCESS.EXE` dépend de la version et de l'installation d'Office). Cette opération réduit la taille du fichier, résout des erreurs inexplicables, et est recommandée périodiquement — en particulier avant de produire un `.accde` (chapitre 20). La routine d'hygiène type consiste à **décompiler, puis compacter/réparer, puis recompiler**. La prudence impose de toujours travailler sur une **copie de sauvegarde** et de fermer les autres instances. Dans une optique de versioning, décompiler et compacter avant l'export garantit que l'on part d'un état propre et cohérent.

## Les outils qui automatisent le va-et-vient

Écrire et maintenir soi-même ses routines d'export/import est possible, mais des **add-ins communautaires** automatisent ce va-et-vient et réduisent considérablement la friction. Le plus abouti aujourd'hui est l'add-in **msaccess-vcs-addin** d'Adam Waller, qui ajoute un ruban dédié pour exporter l'application vers des fichiers source et la reconstruire ensuite, en vue d'un usage avec Git ou GitLab. D'autres outils existent (OASIS-SVN, Ivercy notamment), tous bâtis sur les mêmes méthodes `SaveAsText`/`LoadFromText`.

Un point important : **ces add-ins ne sont pas, à eux seuls, des systèmes de gestion de versions**. Ils exportent et importent ; le versioning proprement dit reste assuré par Git (ou un équivalent), exactement comme pour tout autre projet logiciel. Ils offrent aussi des options précieuses, comme la séparation, pour chaque formulaire ou état, de la **mise en page** et du **code VBA** dans deux fichiers distincts : la revue de code y gagne énormément, car on peut se concentrer sur les changements de code sans être noyé sous le texte de disposition. À l'inverse, l'ancienne intégration via Visual SourceSafe est obsolète et abandonnée depuis longtemps : il ne faut plus s'appuyer dessus.

## Un workflow Git pour Access

L'organisation type d'un dépôt repose sur un principe : **le texte est la source de vérité, le `.accdb` se régénère**. Le dépôt contient un dossier source (les fichiers texte exportés), les routines ou l'add-in d'export/import, la documentation de projet (section 24.3), et un fichier `.gitignore` qui exclut le binaire de travail et les fichiers de verrouillage.

```
# Fichiers binaires et verrous Access (régénérés depuis l'export texte)
*.accdb
*.accde
*.laccdb
*.ldb
~$*.*
```

Le cycle de travail devient alors familier : on développe dans le `.accdb`, on **exporte** vers le texte, on **relit les différences**, puis on **valide** (commit). Pour récupérer le travail d'un collègue, on **tire** (pull) les fichiers texte et on **reconstruit** le `.accdb` à partir d'eux. Les équipes travaillent ainsi chacune sur sa propre copie de la base, exportent et valident leurs changements, qui sont **relus et fusionnés au niveau des fichiers source**, avant que la base ne soit reconstruite pour combiner le tout. Le branchement et la fusion deviennent possibles : le code des modules fusionne très bien, le texte de mise en page des formulaires demande davantage de précaution.

## Les limites à connaître

Soyons honnêtes : ce dispositif n'est pas aussi fluide que Git sur un projet entièrement textuel. Il impose une **discipline** (exporter avant de valider) et une étape d'export/import que les add-ins allègent sans la supprimer. La fusion du texte des **formulaires et états** peut être délicate, là où le code se fusionne aisément. Les **données** ne sont pas versionnées de cette façon — on versionne le schéma, pas le contenu, ce qui est d'ailleurs souhaitable. Enfin, la **co-édition simultanée** d'un même `.accdb` reste difficile : le modèle est « chacun sur sa copie, fusion via Git », parfaitement adapté au code, mais nécessitant de la coordination pour les modifications concurrentes au niveau des formulaires.

## En résumé

Une application Access est un fichier binaire, ce qui rend Git inopérant si on l'y place tel quel : on n'obtient que des sauvegardes, sans différences lisibles ni fusion. La solution consiste à exporter les objets en texte, via les méthodes `SaveAsText`/`LoadFromText`, et à versionner ces fichiers ; la décompilation par `/decompile`, associée au compactage, assure que l'on part d'une source propre et cohérente. Des add-ins communautaires comme celui d'Adam Waller automatisent l'export et l'import, mais le versioning reste assuré par Git lui-même. Le dépôt prend le texte pour source de vérité et régénère le binaire, ce qui ouvre l'historique, la relecture, le branchement et la collaboration — au prix d'une discipline d'export et de quelques limites sur les formulaires et les données. Une fois le code ainsi rendu lisible et comparable, deux pratiques essentielles deviennent enfin praticables : la revue de code et le refactoring, objet de la section suivante.

⏭️ [24.5. Revue de code et refactoring d'applications existantes](/24-bonnes-pratiques-ressources/05-revue-code-refactoring.md)
