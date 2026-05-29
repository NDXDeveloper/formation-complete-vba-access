🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 20.4. Signature numérique des macros et confiance du code

Lorsqu'un utilisateur ouvre une base Access contenant du code VBA ou des macros, la sécurité d'Office n'accorde pas sa confiance automatiquement. Selon les réglages en vigueur, le code peut être purement et simplement désactivé, ou bien l'utilisateur reçoit un avertissement l'invitant à autoriser le contenu. La signature numérique est le mécanisme qui permet de lever cette défiance de façon maîtrisée, en prouvant deux choses : qui a produit le code, et que celui-ci n'a pas été modifié depuis.

Il importe de bien situer cette protection par rapport aux précédentes. Le `.accde` (voir [section 20.3](03-compilation-accde.md)) empêche de lire et de modifier le code ; le mot de passe (voir [section 20.2](02-protection-mot-de-passe.md)) chiffre les données. La signature, elle, ne cache rien et ne chiffre rien : elle ne sert ni à protéger le code ni à protéger les données, mais à établir la confiance dans l'origine et l'intégrité du code. C'est une réponse au problème de la confiance, pas à celui de la confidentialité.

## Ce qu'apporte une signature numérique

Une signature numérique remplit deux fonctions. La première est l'authentification : elle identifie de façon vérifiable l'éditeur du code, c'est-à-dire la personne ou l'organisation qui l'a signé. La seconde est l'intégrité : elle garantit que le code n'a pas été altéré depuis sa signature. Cette garantie d'intégrité a une conséquence directe et parfois déroutante — la moindre modification ultérieure du code invalide la signature. Toute évolution impose donc de signer à nouveau.

En revanche, et c'est à souligner, la signature ne dissimule pas le code et n'en empêche pas la copie. Un projet signé mais non compilé reste parfaitement lisible. Signature et `.accde` répondent à des besoins distincts et se combinent fréquemment.

## Les certificats de signature de code

Signer suppose de disposer d'un certificat de signature de code. La confiance qu'un certificat inspire dépend entièrement de son origine, et l'on distingue trois grandes catégories.

Le certificat auto-signé est créé par soi-même, sans autorité tierce. Il n'est reconnu comme digne de confiance que sur le poste où il a été créé, ou là où on l'a explicitement installé et approuvé. Il convient pour le développement et l'usage personnel, mais en aucun cas pour une distribution à d'autres utilisateurs, chez qui il apparaîtra comme provenant d'un éditeur non approuvé.

Le certificat délivré par une autorité de certification interne s'inscrit dans l'infrastructure à clés publiques d'une entreprise. Diffusé sur les postes du domaine par stratégie de groupe, il permet une confiance à l'échelle de l'organisation, ce qui en fait la solution naturelle pour une application interne.

Le certificat délivré par une autorité de certification commerciale, enfin, est émis par un prestataire reconnu après vérification d'identité. C'est la voie nécessaire pour une distribution large à des utilisateurs externes ; elle est payante.

Dans tous les cas, le mécanisme de confiance repose sur deux maillons. Le certificat doit d'abord remonter, par sa chaîne de certification, à une autorité racine présente dans le magasin des autorités de certification racines de confiance du poste. L'éditeur doit ensuite figurer dans la liste des éditeurs approuvés. Ces deux notions sont gérées dans le centre de gestion de la confidentialité, détaillé en [section 20.5](05-parametres-securite-macro.md).

## Créer un certificat auto-signé avec SELFCERT

Office est livré avec un petit utilitaire permettant de fabriquer un certificat auto-signé, nommé « Certificat numérique pour les projets VBA » (fichier `selfcert.exe`, situé dans le dossier d'installation d'Office). Il suffit d'indiquer un nom, et le certificat est créé dans le magasin personnel de l'utilisateur courant.

Sa limite doit être parfaitement comprise : ce certificat n'est, par défaut, digne de confiance que sur la machine où il a été généré, et encore faut-il l'avoir approuvé. Sur tout autre poste, il sera considéré comme non fiable. Il ne permet donc pas de signer une application destinée à être diffusée. Son usage légitime se cantonne aux tests et au travail du développeur sur son propre poste.

## Signer le projet VBA

La signature s'applique au projet VBA depuis l'éditeur de code (accessible par Alt+F11), via le menu Outils puis Signature numérique. On y sélectionne le certificat à utiliser, puis l'on enregistre la base.

Le moment où l'on signe a son importance. Comme toute modification invalide la signature, il faut signer une fois le code finalisé, juste avant la livraison, et enregistrer dans la foulée. Si vous reprenez ensuite le code, ne serait-ce que pour une correction mineure, la signature tombe et doit être réappliquée.

## Ce qui se passe à l'ouverture d'une base signée

Le comportement à l'ouverture dépend de la confiance accordée à l'éditeur et de la validité de la signature. Si l'éditeur figure déjà parmi les éditeurs approuvés, le code s'exécute sans aucune interruption. Si la signature est valide mais que l'éditeur n'est pas encore approuvé, Office présente à l'utilisateur les informations sur l'éditeur et propose de lui accorder sa confiance pour l'avenir, ce qui l'ajoute à la liste des éditeurs approuvés et rend les ouvertures suivantes silencieuses. En revanche, si la signature est invalide — code modifié, certificat expiré ou non fiable —, un avertissement s'affiche et le code peut être désactivé.

Ce comportement n'est pleinement effectif que si le réglage de sécurité des macros le prévoit, en particulier l'option qui n'autorise que les macros signées par un éditeur approuvé. La configuration correspondante relève de la [section 20.5](05-parametres-securite-macro.md).

Il peut être utile, côté application, de savoir si le code s'exécute effectivement dans un contexte de confiance. La propriété `IsTrusted` de `CurrentProject` renseigne sur ce point :

```vba
If CurrentProject.IsTrusted Then
    ' Le contenu actif (code, macros) est autorisé : tout fonctionnera.
Else
    ' Le code est bloqué par la sécurité : prévenir l'utilisateur
    ' et l'orienter vers un emplacement approuvé ou l'activation du contenu.
    MsgBox "Le contenu de cette application est désactivé par la sécurité.", _
           vbExclamation
End If
```

## Gérer l'expiration et l'horodatage

Un certificat possède une date d'expiration, qu'il faut anticiper. De façon générale en signature de code, l'horodatage par une autorité de confiance permet à une signature de rester valable au-delà de l'expiration du certificat, pour les fichiers signés pendant sa période de validité. Indépendamment de cette possibilité, la bonne pratique consiste à surveiller l'échéance du certificat, à le renouveler en temps voulu et à re-signer l'application avec le certificat renouvelé, afin d'éviter qu'un certificat expiré ne déclenche soudainement des avertissements chez les utilisateurs.

## Articulation avec le format ACCDE et le déploiement

La signature étant appliquée au projet VBA depuis l'éditeur, elle s'effectue nécessairement sur le fichier source `.accdb`, qui seul possède un projet ouvrable dans l'éditeur. Si vous distribuez une version compilée, signez donc le projet sur la source avant de générer le `.accde` (voir [section 20.3](03-compilation-accde.md)), car le projet d'un `.accde` ne peut plus être ouvert dans l'éditeur pour y appliquer une signature après coup.

Dans la pratique, pour un déploiement interne, beaucoup d'équipes préfèrent une approche alternative à la signature : placer l'application dans un emplacement approuvé. Le code s'y exécute alors sans avertissement, indépendamment de toute signature. Le choix entre les deux relève d'un arbitrage. La signature voyage avec le fichier et atteste son origine, ce qui est précieux pour une diffusion externe. L'emplacement approuvé, lui, est un réglage propre à chaque poste, souvent plus simple à mettre en place en interne mais qui ne dit rien de la provenance du fichier. Ces emplacements approuvés sont traités en [section 20.5](05-parametres-securite-macro.md).

## Limites et bonnes pratiques

La principale limite tient à la nature même de la signature : elle ne protège ni le code ni les données, elle ne fait qu'attester d'une origine et d'une intégrité. Un certificat auto-signé ne passe pas à l'échelle de la distribution, et un certificat commercial représente un coût et une démarche de vérification d'identité. Enfin, le caractère fragile de la signature face à toute modification impose de l'intégrer comme une étape finale du processus de livraison, après que le code a été figé, et de la réappliquer à chaque nouvelle version.

## Ce qu'il faut retenir

La signature numérique est l'outil de la confiance, à distinguer nettement des outils de la confidentialité que sont le mot de passe et le `.accde`. Elle authentifie l'éditeur et garantit l'intégrité du code, offrant aux utilisateurs une ouverture sans avertissement dès lors que l'éditeur est approuvé. Elle exige un certificat dont la portée correspond à l'usage visé — auto-signé pour le développement, autorité interne pour l'entreprise, autorité commerciale pour une diffusion externe — et doit être réappliquée à chaque modification du code. Pour un déploiement interne, l'emplacement approuvé constitue souvent une alternative plus simple, dont les réglages, comme ceux de la sécurité des macros, font l'objet de la section suivante.

---


⏭️ [20.5. Paramètres de sécurité macro et centre de gestion de la confidentialité](/20-securite-protection/05-parametres-securite-macro.md)
