🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 21.6. Compatibilité entre versions d'Access (2016, 2019, 2021, M365)

Une application Access déployée tourne rarement sur un parc parfaitement homogène. D'un poste à l'autre, on rencontre des versions différentes — Access 2016, 2019, 2021, ou la version d'abonnement Microsoft 365 — héritées d'installations successives, de politiques d'achat variées ou de cycles de renouvellement décalés. Savoir ce qui reste compatible d'une version à l'autre, et surtout ce qui ne l'est pas, est indispensable pour éviter qu'une application validée sur le poste du développeur ne se comporte de façon imprévisible chez certains utilisateurs.

Cette section fait le point sur le paysage des versions, explique pourquoi le format de données est largement compatible alors que d'autres aspects ne le sont pas, et propose une stratégie de déploiement sur un parc hétérogène.

## Le paysage des versions d'Access

Les versions concernées se répartissent en deux familles. D'un côté, les **versions perpétuelles** — Access 2016, 2019 et 2021 — au jeu de fonctionnalités figé, achetées une fois pour toutes. De l'autre, la version **Microsoft 365**, fournie par abonnement et mise à jour en continu. Microsoft 365 est en permanence à jour des dernières fonctionnalités et disponible sous forme d'abonnement mensuel ou annuel.

Au-delà des quatre versions mentionnées dans le titre de cette section, le paysage compte désormais une version perpétuelle plus récente, Office 2024 (et sa déclinaison Office LTSC 2024), qui reprend une partie des fonctionnalités de Microsoft 365 et est prise en charge jusqu'en octobre 2029. Un parc réel en 2026 peut donc mêler toutes ces versions.

### Cycle de support : l'état actuel

Le cycle de vie de chaque version influence directement la stratégie de déploiement, car maintenir une application sur une version hors support expose à des risques de sécurité et à l'absence de correctifs.

L'état au moment de la rédaction est le suivant :

- **Access 2016 et 2019** : le support d'Office 2016 et d'Office 2019 a pris fin le 14 octobre 2025, ces versions ne reçoivent donc plus ni mises à jour de sécurité ni support technique, même si les applications continuent de fonctionner.
- **Access 2021** : Office LTSC 2021 atteint la fin de support le 13 octobre 2026 ; après cette date, Microsoft ne fournira plus de support technique, de corrections de bugs ni de mises à jour de sécurité. Office 2021 a été lancé le 5 octobre 2021, et sa fin de support survient exactement cinq ans après sa sortie.
- **Microsoft 365** : maintenue en continu, sans date de fin tant que l'abonnement est actif.
- **Office 2024 / LTSC 2024** : version perpétuelle la plus récente, prise en charge jusqu'en octobre 2029.

Concrètement, un parc reposant encore sur Access 2016 ou 2019 devrait être considéré comme en sursis, et la planification d'une montée de version fait partie intégrante de la maintenance de l'application (chapitre 24).

## La bonne nouvelle : un format de fichier stable

L'élément le plus rassurant en matière de compatibilité est le **format de fichier**. Le format `.accdb`, reposant sur le moteur ACE, est stable depuis Access 2007. Toutes les versions concernées — 2016, 2019, 2021, 2024 et Microsoft 365 — utilisent ce même format et le même moteur de base de données. Un fichier `.accdb` créé dans l'une de ces versions s'ouvre dans les autres sans conversion ni perte.

La conséquence pratique est majeure : **la compatibilité au niveau des données est, en pratique, un non-problème** sur ces versions. Le back-end partagé est lisible et modifiable indifféremment par des postes équipés de versions différentes, ce qui rend possible un parc hétérogène travaillant sur les mêmes données. L'ancien format `.mdb` (moteur Jet) reste ouvrable mais relève désormais du legacy, et n'a plus lieu d'être pour une application moderne.

Cette stabilité du format ne signifie cependant pas que tout est compatible : les difficultés se concentrent ailleurs, dans le code, les références et les fonctionnalités.

## Le point sensible : l'ACCDE est lié à la version

Contrairement au `.accdb`, le fichier compilé `.accde` n'est **pas** indépendant de la version. Un ACCDE généré avec une version donnée d'Access peut refuser de s'ouvrir dans une version plus ancienne, comme cela a été établi à la section 21.1.

Pour un déploiement sur parc hétérogène, c'est le point central : il faut **générer l'ACCDE avec la version la plus ancienne présente sur le parc cible**. Si des postes tournent encore sous Access 2016, l'ACCDE doit être produit depuis Access 2016 pour rester ouvrable partout. Produire le livrable depuis une version récente — Microsoft 365, par exemple — le rendrait inutilisable sur les postes équipés d'une version antérieure, alors même que le format de données, lui, serait parfaitement compatible.

## Les références et le late binding

Une cause fréquente d'incompatibilité tient aux **références** vers des bibliothèques externes. Lorsqu'un code utilise une liaison précoce (early binding) vers une version précise d'une bibliothèque — par exemple « Microsoft Excel 16.0 Object Library » pour l'automation Excel (chapitre 22) — cette référence peut échouer à se résoudre sur un poste équipé d'une version différente de la bibliothèque. Une référence manquante provoque une erreur de compilation qui, en VBA, **désactive l'ensemble du projet** : plus aucun code ne s'exécute, même celui qui n'avait rien à voir avec la bibliothèque absente.

Deux parades existent. La première consiste à privilégier la **liaison tardive** (late binding) pour les bibliothèques externes susceptibles de varier d'un poste à l'autre, ce qui supprime la dépendance à une version précise au prix de la perte de l'IntelliSense ; le choix entre liaison précoce et tardive est traité à la section 2.6. La seconde consiste, lorsqu'on conserve la liaison précoce, à référencer la version la plus ancienne présente sur le parc, afin de maximiser la compatibilité ascendante.

## Les fonctionnalités propres aux versions récentes

Chaque nouvelle version d'Access introduit des fonctionnalités, des types de données ou des propriétés qui n'existent pas dans les versions antérieures. Un code qui s'appuie sur une telle nouveauté échouera, voire empêchera l'ouverture correcte du fichier, sur une version qui ne la connaît pas.

L'exemple le plus parlant est le type de données **Grand Nombre** (Large Number / BigInt), introduit avec Access 2019 et Microsoft 365 mais absent d'Access 2016. Une base utilisant un champ de ce type ne peut pas être exploitée correctement sous Access 2016. D'autres nouveautés — graphiques modernes, intégrations Dataverse ou SharePoint récentes, certaines propriétés ajoutées au fil des versions — posent le même type de difficulté.

La règle qui découle du choix de cibler la version la plus ancienne est donc claire : **s'interdire les fonctionnalités absentes de cette version de référence.** Concevoir l'application en se limitant au socle commun à toutes les versions du parc est le moyen le plus sûr de garantir un comportement uniforme.

## Microsoft 365 : une cible mouvante

Microsoft 365 occupe une place particulière car, à la différence des versions perpétuelles au comportement figé, elle est mise à jour mensuellement. Cette évolution continue présente un double visage.

C'est d'un côté un avantage : les utilisateurs disposent des dernières fonctionnalités et des correctifs les plus récents. C'est de l'autre un facteur de risque opérationnel : une mise à jour de Microsoft 365 peut introduire de nouvelles fonctionnalités qui, si elles sont utilisées, briseront la compatibilité avec les versions perpétuelles ; plus délicat encore, une mise à jour peut occasionnellement introduire une **régression** affectant une application jusque-là stable. Une application validée un jour donné peut ainsi se voir perturbée par une évolution ultérieure de Microsoft 365, sans qu'aucune modification n'ait été apportée à l'application elle-même.

La conséquence est une vigilance accrue : sur un parc en Microsoft 365, il faut suivre les évolutions du produit et retester périodiquement l'application, en gardant à l'esprit que la « version » des utilisateurs n'est jamais réellement figée.

## Version et bitness : deux axes distincts

Il importe de ne pas confondre la **version** d'Access (2016, 2019, 2021, 2024, M365) avec son **architecture** 32 ou 64 bits. Ce sont deux dimensions indépendantes : chaque version existe en 32 et en 64 bits. Pour qu'un ACCDE soit utilisable sur un poste, il faut que **la version et la bitness** correspondent toutes deux. Un parc peut donc être hétérogène sur les deux axes à la fois, ce qui complexifie d'autant le packaging. La problématique 32/64 bits, et notamment son impact sur les déclarations d'API, fait l'objet de la section 21.7.

## Identifier la version d'exécution par code

Il est parfois utile de connaître, à l'exécution, la version d'Access sur laquelle tourne l'application. La fonction `SysCmd` avec la constante `acSysCmdAccessVer` renvoie le numéro de version interne :

```vba
' Renvoie la version interne d'Access
Public Function VersionAccess() As String
    VersionAccess = SysCmd(acSysCmdAccessVer)
End Function
```

La propriété `Application.Version` fournit la même information. Mais un piège majeur attend ici le développeur : **Access 2016, 2019, 2021, 2024 et Microsoft 365 partagent tous le même numéro de version majeure, « 16.0 ».** La fonction `acSysCmdAccessVer` comme `Application.Version` renvoient donc « 16.0 » pour l'ensemble de ces versions, et **ne permettent pas de les distinguer**. (Seules les versions plus anciennes diffèrent : « 15.0 » pour 2013, « 14.0 » pour 2010.)

Pour identifier précisément le produit — distinguer un poste 2016 d'un poste Microsoft 365, par exemple — il faut descendre au niveau du **numéro de build**, en inspectant le registre Windows ou la version du fichier exécutable. La lecture du registre est traitée à la section 22.9. En pratique, fonder une logique applicative sur la distinction fine des versions doit rester exceptionnel : il est presque toujours préférable de viser le socle commun plutôt que d'aiguiller le code selon la version détectée.

## Stratégie de déploiement multi-versions

La gestion d'un parc hétérogène se ramène à quelques principes simples et cohérents :

- **Identifier la version la plus ancienne** présente sur le parc cible. Elle devient la version de référence.
- **Développer et valider en visant ce socle**, en s'interdisant les fonctionnalités qui lui sont postérieures.
- **Générer l'ACCDE avec cette version de référence**, et dans chaque bitness nécessaire (section 21.7).
- **Privilégier la liaison tardive** pour les automations externes, afin d'échapper aux ruptures de références entre versions (section 2.6).
- **Tester sur chaque version réellement présente** sur le parc, idéalement au moyen de machines virtuelles dédiées, plutôt que de présumer qu'un fonctionnement correct sur une version vaut pour toutes.

À plus long terme, un parc durablement hétérogène ou la persistance de versions hors support sont autant de signaux invitant à envisager une évolution d'architecture, voire une migration vers un socle plus robuste (chapitre 23).

## Articulation avec le reste du déploiement

La compatibilité entre versions recoupe plusieurs autres sujets du chapitre et de la formation :

- Le caractère version-sensible de l'**ACCDE** a été établi lors du packaging (section 21.1).
- La dimension **32/64 bits**, distincte de la version, est traitée à la section 21.7.
- Le choix **early binding / late binding** relève de la section 2.6.
- L'identification fine d'une version par le **registre** est abordée à la section 22.9.
- La **migration** vers un socle plus pérenne, lorsque l'hétérogénéité devient ingérable, fait l'objet du chapitre 23.

## Points de vigilance

Plusieurs erreurs reviennent fréquemment en contexte multi-versions :

- **Générer l'ACCDE depuis une version trop récente** par rapport au parc, le rendant inutilisable sur les postes plus anciens, alors que le format de données serait pourtant compatible.
- **Confondre compatibilité des données et compatibilité du code** : le `.accdb` traverse les versions sans peine, mais l'ACCDE, les références et les fonctionnalités, non.
- **Utiliser une fonctionnalité récente** (type Grand Nombre, graphiques modernes) qui exclut de fait les versions antérieures du parc.
- **S'appuyer sur `Application.Version` pour distinguer 2016, 2021 ou Microsoft 365**, alors que toutes renvoient « 16.0 ».
- **Oublier la nature mouvante de Microsoft 365** et ne pas retester l'application après les mises à jour du produit.
- **Maintenir une application sur une version hors support** (2016, 2019), au mépris des risques de sécurité et d'absence de correctifs.

⏭️ [21.7. Différences 32 bits / 64 bits — déclarations API et compatibilité](/21-deploiement-distribution/07-32-bits-vs-64-bits.md)
