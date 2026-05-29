🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 20. Sécurité et protection

La sécurité d'une application Access est un sujet aussi important qu'il est souvent mal compris. Beaucoup de développeurs considèrent Access comme un outil bureautique « léger » et négligent la protection de leurs bases, jusqu'au jour où un fichier `.accdb` est copié sur une clé USB, où le code source d'une application est révélé à un concurrent, ou bien où des données confidentielles sont modifiées sans qu'aucune trace ne permette de savoir qui, quand et quoi. À l'inverse, certains surestiment les garanties offertes par Access et croient à tort qu'un simple mot de passe suffit à rendre une base inviolable.

Ce chapitre a pour objectif de donner une vision à la fois réaliste et complète des mécanismes de protection disponibles dans Access : ce qu'ils protègent réellement, leurs limites, et surtout la manière de les combiner pour bâtir une application raisonnablement sûre, adaptée au contexte dans lequel elle est déployée.

## Une mise en garde indispensable

Il faut le dire d'emblée, sans détour : Access n'est pas une plateforme de sécurité de niveau entreprise. Le moteur ACE, qui fait fonctionner les bases `.accdb`, n'a jamais été conçu pour résister à un attaquant motivé disposant d'un accès physique au fichier. La plupart des protections présentées dans ce chapitre relèvent davantage de la dissuasion et de la maîtrise des risques courants (erreur humaine, curiosité d'un utilisateur, copie accidentelle) que d'une protection absolue contre une attaque ciblée.

Le principe directeur à retenir est simple : dès lors que des données sont véritablement sensibles ou que l'enjeu de confidentialité est élevé, la bonne réponse n'est pas d'empiler des protections Access mais de migrer le stockage vers un véritable serveur de bases de données comme SQL Server, qui offre une authentification robuste, un chiffrement éprouvé et une gestion fine des permissions. Cette piste est traitée au [chapitre 23](../23-migration-interoperabilite/README.md). Le présent chapitre se concentre sur ce qu'il est raisonnable et possible de faire à l'intérieur de l'écosystème Access lui-même.

## Une protection en couches

Plutôt que de chercher une mesure unique et miraculeuse — qui n'existe pas —, la démarche recommandée consiste à superposer plusieurs lignes de défense complémentaires, chacune couvrant un aspect différent. C'est le principe de la défense en profondeur. Ce chapitre est organisé selon ces différentes couches.

La première est la couche du fichier lui-même : le mot de passe de base de données et le chiffrement qui l'accompagne visent à empêcher l'ouverture du fichier et la lecture brute de son contenu. La deuxième est la couche du code : la compilation en `.accde`, la signature numérique et les paramètres du centre de gestion de la confidentialité protègent la propriété intellectuelle (votre code VBA) et la confiance accordée à l'application au démarrage. La troisième est la couche applicative : une gestion des droits par table de rôles et un dispositif d'audit permettent de contrôler ce que chaque utilisateur peut faire et de conserver une trace de ses actions. Enfin, la couche des données elle-même peut être renforcée par un chiffrement applicatif des informations les plus sensibles, indépendant du chiffrement global du fichier.

Aucune de ces couches n'est suffisante isolément, mais leur combinaison réfléchie, calibrée selon le niveau de risque réel, constitue une stratégie de sécurité cohérente.

## Objectifs du chapitre

À l'issue de ce chapitre, vous serez en mesure de :

- Comprendre ce que le modèle de sécurité d'Access protège réellement, et ce qu'il ne protège pas, en tenant compte de son évolution historique.
- Mettre en place les protections natives du fichier (mot de passe, chiffrement) et en connaître les limites.
- Protéger votre code VBA contre la lecture et la modification grâce à la compilation en `.accde`.
- Maîtriser la signature numérique et les réglages de sécurité des macros pour éviter les blocages et les avertissements intempestifs au démarrage.
- Concevoir une gestion applicative des droits utilisateurs fondée sur une table de rôles.
- Mettre en œuvre un mécanisme d'audit des accès et des modifications par code.
- Chiffrer les données sensibles stockées dans les tables lorsque le contexte l'exige.

## Prérequis

Ce chapitre suppose une bonne familiarité avec le modèle objet d'Access et l'objet `DoCmd` (chapitres 4 et 5), avec la manipulation des données via DAO (chapitre 9) et avec le SQL exécuté par code (chapitre 11), notamment pour la gestion des droits et l'audit. La gestion des erreurs (chapitre 13) sera également mobilisée, car les routines de sécurité doivent être particulièrement robustes. Une connaissance des formats de fichiers Access (chapitre 1) facilitera la compréhension des sections sur le mot de passe et la compilation.

## Plan du chapitre

- [20.1. Modèle de sécurité Access — historique et situation actuelle](01-modele-securite-access.md)
  Pourquoi l'ancienne sécurité au niveau utilisateur a disparu avec les fichiers `.accdb`, et quel modèle la remplace aujourd'hui.

- [20.2. Protection par mot de passe de la base de données](02-protection-mot-de-passe.md)
  Poser, modifier et retirer un mot de passe, comprendre le chiffrement associé et ses limites pratiques.

- [20.3. Compilation en ACCDE — verrouillage du code VBA](03-compilation-accde.md)
  Générer un fichier `.accde` pour empêcher la lecture et la modification du code, et organiser le couple source/exécutable.

- [20.4. Signature numérique des macros et confiance du code](04-signature-numerique.md)
  Signer son projet avec un certificat pour rassurer Access et les utilisateurs sur l'origine du code.

- [20.5. Paramètres de sécurité macro et centre de gestion de la confidentialité](05-parametres-securite-macro.md)
  Comprendre les emplacements approuvés, les niveaux de sécurité des macros et la barre des messages.

- [20.6. Gestion applicative des droits utilisateurs (table de rôles)](06-gestion-droits-utilisateurs.md)
  Construire un système de rôles et de permissions piloté par des tables et appliqué dans l'interface par code.

- [20.7. Audit des accès et des modifications par code](07-audit-acces-modifications.md)
  Journaliser les connexions et tracer les créations, modifications et suppressions d'enregistrements.

- [20.8. Chiffrement des données sensibles dans les tables](08-chiffrement-donnees-sensibles.md)
  Protéger des champs précis (mots de passe applicatifs, données personnelles) par un chiffrement au niveau des données.

## Fil conducteur

Tout au long de ce chapitre, gardez en tête la question essentielle : de quoi cherche-t-on à se protéger, et à quel coût ? La sécurité parfaite n'existe pas, et un excès de protections mal calibrées nuit à l'utilisabilité de l'application sans apporter de garantie réelle. L'objectif n'est pas de tout verrouiller, mais de placer le bon niveau de protection au bon endroit, en proportion du risque effectivement encouru.

---


⏭️ [20.1. Modèle de sécurité Access — historique et situation actuelle](/20-securite-protection/01-modele-securite-access.md)
