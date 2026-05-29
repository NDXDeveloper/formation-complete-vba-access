🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 20.5. Paramètres de sécurité macro et centre de gestion de la confidentialité

Les sections précédentes ont évoqué à plusieurs reprises la « confiance » accordée au code et les avertissements qui apparaissent à l'ouverture d'une base. C'est ici que tout se décide. Le centre de gestion de la confidentialité (en anglais *Trust Center*) est l'endroit où Office et Access déterminent si le code VBA et les macros d'une base ont le droit de s'exécuter, et dans quelles conditions. Comprendre ces réglages est indispensable, car une application parfaitement écrite restera muette si la sécurité bloque son code.

On accède au centre de gestion de la confidentialité depuis l'onglet Fichier, puis Options, rubrique Centre de gestion de la confidentialité, et enfin le bouton Paramètres du Centre de gestion de la confidentialité. C'est un réglage propre à chaque poste et à chaque utilisateur, ce qui a des conséquences importantes pour le déploiement, abordées plus loin.

## La barre des messages

La manifestation la plus visible de ces réglages est la barre des messages. Lorsque le code d'une base est désactivé par la sécurité, une bande jaune apparaît sous le ruban, signalant un avertissement de sécurité et indiquant qu'une partie du contenu actif a été désactivée. Elle comporte un bouton Activer le contenu. En cliquant dessus, l'utilisateur autorise le code pour ce fichier et, ce faisant, l'enregistre comme document approuvé, si bien que l'avertissement ne réapparaîtra plus pour ce fichier précis. C'est l'interaction de base entre l'utilisateur et la sécurité d'Access, mais elle suppose que l'utilisateur sache et accepte de cliquer, ce qui n'est pas toujours souhaitable dans une application livrée.

## Les paramètres des macros

Le cœur du réglage se trouve dans la section Paramètres des macros du centre de gestion de la confidentialité. Il faut noter d'emblée que le terme « macros » y est employé au sens large : il englobe non seulement les macros Access, mais aussi le code VBA et, plus généralement, le contenu actif. Quatre options s'excluent mutuellement.

La première, Désactiver toutes les macros sans notification, bloque tout code sans même prévenir l'utilisateur : aucune barre des messages n'apparaît. C'est le réglage le plus restrictif. La deuxième, Désactiver toutes les macros avec notification, est le réglage par défaut : le code est désactivé mais la barre des messages s'affiche, laissant à l'utilisateur la possibilité d'activer le contenu. La troisième, Désactiver toutes les macros à l'exception des macros signées numériquement, n'autorise que le code signé par un éditeur approuvé, ce qui donne tout son sens à la signature numérique vue en [section 20.4](04-signature-numerique.md). La quatrième, Activer toutes les macros, exécute tout code sans contrôle ; elle est explicitement déconseillée, car elle ouvre une brèche à l'échelle de la machine pour n'importe quelle base ouverte ensuite.

## Le contenu désactivé : ce que cela implique réellement

Lorsqu'une base est dans l'état « contenu désactivé », les conséquences vont au-delà d'un simple avertissement. Le code VBA ne s'exécute pas, ce qui signifie que les procédures événementielles des formulaires et des états ne se déclenchent pas, que les fonctions personnalisées ne répondent pas, et qu'une application reposant sur du code peut tout simplement paraître non fonctionnelle. Par ailleurs, Access applique en mode non approuvé un cloisonnement qui bloque certaines fonctions et actions jugées potentiellement dangereuses, héritage du « mode bac à sable » (*sandbox mode*) : certaines expressions dans les requêtes ou les contrôles peuvent ainsi être refusées. Il est donc essentiel de comprendre qu'une base non approuvée n'est pas seulement « avec un avertissement », mais réellement bridée dans son fonctionnement. Pour détecter cet état par code, on peut s'appuyer sur la propriété `CurrentProject.IsTrusted`, illustrée en [section 20.4](04-signature-numerique.md).

## Les emplacements approuvés

C'est ici qu'intervient le mécanisme le plus utile pour le déploiement : les emplacements approuvés. Un emplacement approuvé est un dossier désigné comme digne de confiance ; toute base ouverte depuis ce dossier voit son code s'exécuter sans aucun avertissement, indépendamment des paramètres de macros et de toute signature.

La configuration se fait dans la section Emplacements approuvés du centre de gestion de la confidentialité, où l'on ajoute un nouvel emplacement en désignant un dossier. Une case permet d'étendre la confiance aux sous-dossiers. Une autre, Autoriser les emplacements approuvés sur mon réseau, est désactivée par défaut et explicitement présentée comme non recommandée, car faire confiance à un dossier réseau revient à faire confiance à tout ce que d'autres pourraient y déposer.

Pour une application Access interne, la pratique la plus courante consiste donc à installer le front-end dans un dossier local approuvé sur chaque poste. Cela offre une expérience fluide, sans avertissement, sans dépendre d'une signature. La contrepartie en matière de sécurité doit être assumée : tout fichier placé dans ce dossier sera exécuté en confiance, il faut donc éviter de rendre approuvé un dossier partagé ou trop large.

## Les éditeurs approuvés

La section Éditeurs approuvés répertorie les éditeurs, identifiés par leur certificat, dont le code signé est jugé fiable. Cette liste se peuple notamment lorsqu'un utilisateur, confronté à une base signée par un éditeur encore inconnu mais valide, choisit de lui accorder sa confiance. C'est le pendant direct de la signature numérique : sans éditeur approuvé, l'option « n'autoriser que les macros signées » ne laisserait rien passer. Le lien avec la signature est traité en [section 20.4](04-signature-numerique.md).

## Les documents approuvés

La section Documents approuvés mémorise, fichier par fichier, les bases pour lesquelles l'utilisateur a cliqué sur Activer le contenu. Une fois un fichier ainsi approuvé, son code s'exécute sans avertissement lors des ouvertures suivantes. Cette confiance est propre à chaque utilisateur et à chaque fichier précis ; elle peut être effacée, et l'on peut désactiver complètement le mécanisme si l'on souhaite que chaque ouverture redemande l'autorisation. C'est une commodité, mais elle ne remplace pas une stratégie de déploiement fondée sur les emplacements approuvés ou la signature.

## Configuration à grande échelle

Régler manuellement le centre de gestion de la confidentialité sur chaque poste est impraticable au-delà de quelques machines. Ces paramètres étant stockés dans le registre Windows de l'utilisateur, sous une clé propre à Access et à sa version, ils peuvent être déployés de façon centralisée. En entreprise, on recourt typiquement aux stratégies de groupe, qui poussent les emplacements approuvés et les options de macros sur l'ensemble du parc. Cette approche relève du déploiement, abordé au [chapitre 21](../21-deploiement-distribution/README.md), et la manipulation du registre fait l'objet de la [section 22.9](../22-api-windows-integration-office/09-registre-windows.md).

## Et par code ?

Il est techniquement possible de lire, voire d'écrire, ces réglages via le registre, puisqu'ils y sont stockés. Cela ne doit cependant pas conduire une application à abaisser elle-même la sécurité de son hôte, par exemple en s'ajoutant silencieusement comme emplacement approuvé ou en désactivant les contrôles de macros : c'est une mauvaise pratique, contraire à l'esprit du modèle de sécurité, et souvent entravée par Office. La bonne voie reste la configuration documentée, réalisée par un administrateur ou par l'installateur de l'application, ou bien la signature.

Un réglage distinct mérite toutefois d'être signalé, car il prête à confusion : la propriété `Application.AutomationSecurity`. Elle ne concerne pas l'ouverture normale d'une base par un utilisateur, mais le cas où votre code pilote une autre application Office en automation (voir [chapitre 22](../22-api-windows-integration-office/README.md)) et ouvre des documents par programme. Elle permet alors de forcer le comportement de sécurité des macros pour ces documents ouverts par code.

```vba
' Contexte d'AUTOMATION uniquement (par exemple Access ouvrant un classeur Excel).
' Nécessite une référence à la bibliothèque Microsoft Office.
Application.AutomationSecurity = msoAutomationSecurityForceDisable
' force la désactivation des macros dans les documents ouverts par programme,
' afin de ne pas exécuter de code inconnu à leur ouverture
```

## Quelle option choisir ?

Le bon réglage dépend du contexte. Pour une application interne déployée dans un environnement maîtrisé, l'emplacement approuvé local est la solution la plus simple et la plus fiable : aucun avertissement, aucune dépendance à un certificat. Pour une diffusion plus large, en particulier à l'extérieur, la signature numérique combinée à l'approbation de l'éditeur (ou à un emplacement approuvé poussé par l'installateur) atteste l'origine du code tout en évitant les avertissements. Dans tous les cas, on évite l'option Activer toutes les macros, qui supprime indistinctement toute protection pour l'ensemble des bases, et l'on s'abstient d'affaiblir la sécurité depuis l'application elle-même.

## Ce qu'il faut retenir

Le centre de gestion de la confidentialité est l'arbitre qui décide de l'exécution du code dans Access. Son réglage par défaut désactive le code tout en affichant une barre des messages permettant de l'activer. Pour qu'une application livrée fonctionne sans heurt, deux voies principales s'offrent au développeur : l'emplacement approuvé, idéal en interne et déployable par stratégie de groupe, et la signature numérique associée à un éditeur approuvé, adaptée à une diffusion externe. Comprendre que ces paramètres sont propres à chaque poste, et qu'un contenu non approuvé est réellement bridé et non simplement signalé, est la clé pour livrer des applications Access qui s'exécutent comme prévu.

---


⏭️ [20.6. Gestion applicative des droits utilisateurs (table de rôles)](/20-securite-protection/06-gestion-droits-utilisateurs.md)
