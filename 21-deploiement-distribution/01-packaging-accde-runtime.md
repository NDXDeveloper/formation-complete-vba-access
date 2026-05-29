🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 21.1. Packaging d'une application Access (ACCDE, runtime)

Le packaging est la première étape concrète du déploiement : il consiste à transformer le projet ouvert dans votre environnement de développement en un livrable destiné à être installé et exécuté par d'autres. Cette transformation poursuit deux objectifs distincts mais complémentaires : d'une part **protéger l'application** en empêchant la modification du code et de la conception, d'autre part **contrôler l'environnement d'exécution** pour que l'utilisateur final n'ait accès qu'à ce que vous avez prévu. Le format ACCDE répond au premier objectif, le mode runtime au second.

Copier simplement le fichier `.accdb` sur le poste de l'utilisateur n'est pas un packaging : c'est distribuer votre code source. L'utilisateur peut alors ouvrir vos formulaires en mode Création, lire et modifier votre code VBA, altérer la structure des objets — accidentellement ou non — et compromettre la stabilité de l'ensemble. Le packaging consiste précisément à fermer ces portes.

## Qu'est-ce que le packaging d'une application Access ?

Packager une application, c'est produire une version figée et autonome du front-end, prête à être déployée. Cela implique de prendre plusieurs décisions :

- Sous quel format distribuer le front-end : `.accdb` modifiable, ou `.accde` compilé et verrouillé.
- Dans quel mode l'application doit s'ouvrir chez l'utilisateur : mode complet ou mode runtime.
- Comment l'application sera réellement installée sur les postes (sujet traité à la section 21.8).

Ces choix dépendent du contexte : niveau de confiance accordé aux utilisateurs, besoin de protéger une logique métier, présence ou non d'une licence Access complète sur les postes cibles. Le format ACCDE et le mode runtime constituent les deux briques fondamentales sur lesquelles repose toute stratégie de packaging.

## Le format ACCDE : principe et objectifs

Un fichier ACCDE est la version compilée et verrouillée d'un fichier `.accdb`. Lors de sa génération, Access compile l'intégralité du code VBA, en retire le code source éditable pour ne conserver que le pseudo-code compilé (le *p-code*), puis verrouille la conception des objets d'interface. Le résultat est une application qui s'exécute exactement comme l'originale, mais qui ne peut plus être ni inspectée ni modifiée dans ses parties protégées.

La compilation en ACCDE comme mesure de sécurité et de protection du code VBA est traitée en détail à la section 20.3. Ici, l'ACCDE est envisagé sous l'angle du packaging : c'est le format de livrable qui garantit l'intégrité de l'application déployée et empêche toute dérive de conception en production.

Le tableau de correspondance des extensions est le suivant :

| Format source | Format compilé | Format mode runtime |
|---------------|----------------|---------------------|
| `.accdb` (Access 2007+) | `.accde` | `.accdr` |
| `.mdb` (ancien format) | `.mde` | — |

### Ce que l'ACCDE verrouille, et ce qu'il ne verrouille pas

Il est essentiel de bien comprendre la portée exacte de la protection apportée par l'ACCDE, car elle est souvent mal interprétée.

L'ACCDE **empêche** l'utilisateur de : afficher ou modifier le code VBA ; ouvrir les formulaires et les états en mode Création ou en mode Page ; créer de nouveaux formulaires, états ou modules ; importer ou exporter des formulaires, états et modules vers une autre base ; ajouter, supprimer ou modifier les références aux bibliothèques d'objets.

L'ACCDE **n'empêche pas** l'utilisateur de : ouvrir et utiliser normalement les formulaires et états ; consulter, saisir et modifier les données ; modifier la structure des tables et des requêtes, ainsi que les macros, qui ne sont pas protégées par ce format.

Surtout, l'ACCDE **ne chiffre pas les données**. La protection porte uniquement sur le code et la conception de l'interface, non sur le contenu des tables. Pour protéger les données elles-mêmes, il faut recourir au chiffrement par mot de passe de la base (section 20.2), généralement appliqué au fichier back-end. Confondre ces deux mécanismes est une erreur classique : un ACCDE protège votre travail de développeur, pas les informations confidentielles de vos utilisateurs.

## Créer un fichier ACCDE

La génération d'un ACCDE est une opération native d'Access qui ne requiert aucun outil externe, mais elle exige que certaines conditions soient réunies au préalable.

### Prérequis à la génération

Trois conditions doivent être satisfaites pour qu'Access accepte de produire l'ACCDE :

- **Le projet VBA doit compiler sans erreur.** Lancez systématiquement *Débogage > Compiler* dans l'éditeur VBA avant toute génération. La moindre erreur de compilation bloque la création du fichier.
- **Toutes les références doivent être valides.** Une référence manquante ou cassée (bibliothèque COM absente du poste de développement) empêche la compilation. Ce point renvoie aux questions de liaisons précoces et tardives abordées au chapitre 2.
- **La base doit être au format actuel.** L'ACCDE se génère à partir d'un `.accdb`. Un ancien fichier `.mdb` doit d'abord être converti, et produira alors un `.mde`.

Il est également recommandé de travailler sur une copie : on génère l'ACCDE à partir d'un `.accdb` que l'on conserve précieusement, pour les raisons exposées plus bas.

### La procédure pas à pas

Dans les versions modernes d'Access, la génération s'effectue ainsi :

1. Ouvrir le fichier `.accdb` source à partir duquel vous voulez produire le livrable.
2. Vérifier la compilation du projet VBA (*Débogage > Compiler* dans l'éditeur, accessible par `Alt + F11`).
3. Aller dans **Fichier > Enregistrer sous**.
4. Sous **Enregistrer la base de données sous**, sélectionner **Créer ACCDE**.
5. Cliquer sur le bouton **Enregistrer sous**, choisir l'emplacement et le nom du fichier `.accde`.

Access compile alors le projet et produit un nouveau fichier `.accde` distinct de l'original, ce dernier restant inchangé.

## Conserver impérativement le fichier source ACCDB

Un fichier ACCDE est irréversible : il est impossible de le reconvertir en `.accdb` pour récupérer le code source. La suppression du code lors de la compilation est définitive, ce qui constitue précisément la garantie de protection.

La conséquence est une règle absolue de gestion : **le fichier `.accdb` source est le véritable maître de l'application, et l'ACCDE n'en est qu'un dérivé jetable.** Toute évolution future — correction d'un bug, ajout d'une fonctionnalité — se fait sur le `.accdb`, à partir duquel on régénère ensuite un nouvel ACCDE. Perdre le `.accdb` revient à perdre la capacité de faire évoluer l'application.

Dans la pratique, cela impose une organisation rigoureuse :

- Le `.accdb` source est versionné et sauvegardé (les techniques de versioning Git sont abordées à la section 24.4).
- L'ACCDE est considéré comme un artefact de build, régénéré à chaque livraison, et jamais comme un fichier que l'on modifie directement.
- On veille à ne jamais distribuer le `.accdb` par erreur à la place de l'ACCDE.

## Contraintes de compatibilité de l'ACCDE

Le fichier ACCDE, parce qu'il contient du code déjà compilé, est plus sensible que le `.accdb` aux différences d'environnement. Deux dimensions de compatibilité doivent être anticipées dès le packaging.

### Compatibilité entre versions d'Access

Un ACCDE généré avec une version donnée d'Access peut ne pas s'ouvrir dans une version plus ancienne. La règle de prudence consiste donc à **générer l'ACCDE avec la version la plus ancienne d'Access que vous devez prendre en charge sur le parc cible.** Si vos utilisateurs disposent d'un mélange d'Access 2016 et de Microsoft 365, produire l'ACCDE depuis Access 2016 maximise les chances de compatibilité. La gestion fine de la compatibilité entre versions fait l'objet de la section 21.6.

### Architecture 32 bits / 64 bits

C'est une contrainte fréquemment négligée et pourtant déterminante : **un ACCDE généré avec une version 32 bits d'Access ne peut pas être ouvert par une version 64 bits, et inversement.** Le code étant compilé pour une architecture donnée, la portabilité entre bitness n'est pas assurée.

Si votre parc mélange des installations 32 et 64 bits, vous devrez produire et distribuer deux ACCDE distincts, chacun généré avec la bitness correspondante. Cette contrainte est d'autant plus sensible que les déclarations d'API Windows diffèrent elles aussi selon l'architecture (mots-clés `PtrSafe` et `LongPtr`), sujet détaillé à la section 21.7 et au chapitre 22. L'architecture 32/64 bits influence donc à la fois la génération de l'ACCDE et la portabilité du code qu'il contient.

## Le mode runtime : contrôler l'environnement d'exécution

Le second volet du packaging concerne le mode d'exécution. Le **mode runtime** est un état d'Access dans lequel toutes les fonctions de conception et de développement sont masquées : le volet de navigation disparaît, les rubans de conception sont absents, les raccourcis permettant d'accéder à la structure des objets sont désactivés. L'utilisateur ne voit plus que l'application elle-même — ses formulaires, ses menus, ses états — sans aucune trace de l'environnement Access sous-jacent.

Il faut bien distinguer deux notions souvent confondues sous le terme « runtime » :

- **Le mode runtime** est un *mode d'affichage et de comportement* d'Access. C'est l'objet de la présente section.
- **Le moteur Access Runtime** est une *version gratuite et allégée d'Access*, destinée aux postes ne disposant pas de licence Office complète. Son téléchargement, son installation et son rôle dans le déploiement font l'objet de la section 21.2.

Le point clé est que le mode runtime peut être activé indépendamment du moteur Runtime : une installation complète d'Access peut elle aussi ouvrir une application en mode runtime.

### Forcer le mode runtime sans le moteur Runtime

Deux moyens permettent de forcer le mode runtime, que l'utilisateur dispose d'Access complet ou du seul moteur Runtime.

Le premier est l'**extension `.accdr`**. Un fichier `.accdr` n'est rien d'autre qu'un `.accdb` (ou la copie d'un `.accde` renommée) dont l'extension a été changée. Cette extension indique à Access d'ouvrir le fichier directement en mode runtime. C'est la méthode la plus simple pour distribuer une application qui s'ouvrira toujours dans un environnement verrouillé. À noter que cette approche, à elle seule, ne protège pas le code : pour combiner protection et mode runtime, on renomme une copie de l'ACCDE en `.accdr`.

Le second est le **commutateur de ligne de commande `/runtime`**. Il permet de lancer n'importe quelle base en mode runtime à partir d'un raccourci ou d'un script :

```
"C:\Program Files\Microsoft Office\root\Office16\MSACCESS.EXE" "C:\App\MonApp.accde" /runtime
```

Cette technique est particulièrement utile lorsque l'on crée un raccourci de démarrage pour l'utilisateur ou que l'on automatise le lancement de l'application (section 21.8).

### Détecter le mode runtime par code

Il est fréquemment nécessaire d'adapter le comportement de l'application selon qu'elle s'exécute en mode runtime ou en mode complet — par exemple, désactiver certaines fonctions d'administration réservées au développeur, ou ajuster l'interface. La fonction `SysCmd` avec la constante `acSysCmdRuntime` renvoie `True` lorsque l'application s'exécute en mode runtime :

```vba
Public Function EstEnModeRuntime() As Boolean
    ' Renvoie True si l'application s'exécute en mode runtime
    EstEnModeRuntime = (SysCmd(acSysCmdRuntime) = True)
End Function
```

Cette fonction permet ensuite de conditionner certains comportements lors du démarrage ou dans les formulaires :

```vba
If EstEnModeRuntime() Then
    ' Mode runtime : interface verrouillée pour l'utilisateur final
    ' On masque les outils d'administration et de conception applicative
Else
    ' Mode complet : Access installé en version complète
    ' On peut exposer des fonctions de maintenance réservées
End If
```

Tester ce comportement pendant le développement est facilité par le commutateur `/runtime`, qui permet de simuler le mode runtime sans installer le moteur Runtime sur le poste. Le verrouillage complet de l'interface au démarrage (options de démarrage, masquage du volet de navigation) est par ailleurs traité au chapitre 17.

## ACCDE et runtime : deux dimensions complémentaires

Les deux mécanismes répondent à des préoccupations différentes et se combinent dans la plupart des déploiements professionnels :

- **L'ACCDE protège l'application** : il empêche la lecture du code et la modification de la conception. Il porte sur le *contenu* du livrable.
- **Le mode runtime contrôle l'expérience** : il masque l'environnement de développement et restreint l'utilisateur à l'application. Il porte sur le *contexte d'exécution*.

Un packaging complet consiste donc typiquement à générer un ACCDE, puis à le distribuer de telle sorte qu'il s'ouvre en mode runtime — soit par une copie renommée en `.accdr`, soit par un raccourci utilisant le commutateur `/runtime`. On obtient ainsi une application à la fois protégée et présentée dans un environnement épuré.

Ces deux briques ne suffisent toutefois pas à elles seules à un déploiement réel. Elles s'inscrivent dans une chaîne plus large :

- La mise à disposition du moteur d'exécution sur les postes sans licence complète relève de l'**Access Runtime**, détaillé à la section 21.2.
- La diffusion des nouvelles versions du front-end aux utilisateurs est traitée par les **stratégies d'auto-update** (section 21.3).
- L'installation effective sur les postes, raccourcis et paramètres de lancement compris, fait l'objet de l'**installation automatisée** (section 21.8).

## Choisir sa stratégie de packaging

Le choix du format et du mode dépend du contexte de déploiement. Quelques repères :

Distribuer un `.accdb` non compilé ne se justifie que dans un cadre de confiance total, par exemple pour un usage personnel ou un déploiement auprès d'autres développeurs qui doivent pouvoir intervenir sur le code. Dans tous les autres cas, et en particulier dès qu'une application est livrée à des utilisateurs finaux, l'ACCDE s'impose pour garantir l'intégrité du livrable et protéger la logique métier.

Le mode runtime est presque toujours souhaitable en production : il offre une expérience plus propre et plus sûre, et il est de toute façon imposé lorsque l'application tourne sur le moteur Runtime gratuit. Le réserver aux seuls postes équipés du Runtime serait une erreur ; il gagne à être généralisé à l'ensemble du parc.

## Points de vigilance

Plusieurs erreurs récurrentes méritent d'être anticipées dès la phase de packaging :

- **Confondre protection du code et protection des données.** L'ACCDE ne chiffre rien ; les données sensibles exigent un chiffrement distinct (section 20.2).
- **Distribuer le `.accdb` source par mégarde.** Mettre en place une procédure de build claire qui sépare nettement les fichiers sources des artefacts livrables évite cette confusion.
- **Oublier la contrainte de bitness.** Un ACCDE 32 bits ne s'ouvrira pas sur un poste Access 64 bits ; vérifier l'architecture du parc cible avant de générer le livrable.
- **Générer l'ACCDE depuis une version d'Access trop récente** par rapport au parc cible, le rendant inutilisable sur les postes équipés d'une version antérieure.
- **Négliger le mode runtime au développement.** Tester l'application avec le commutateur `/runtime` révèle des comportements (volet de navigation absent, fonctions de conception indisponibles) qui n'apparaissent jamais en mode complet et qui, non anticipés, se manifestent uniquement chez l'utilisateur.

⏭️ [21.2. Access Runtime — déploiement sans licence complète](/21-deploiement-distribution/02-access-runtime.md)
