🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 20.3. Compilation en ACCDE — verrouillage du code VBA

Le mot de passe de base de données vu en [section 20.2](02-protection-mot-de-passe.md) protège les données au repos, mais il ne protège pas le travail du développeur : une fois la base ouverte, le code VBA, les formulaires et les états restent visibles et modifiables. Pour verrouiller ce contenu, Access propose un format dédié, le `.accde`. Il s'agit de la méthode robuste pour protéger sa propriété intellectuelle, par opposition au mot de passe de projet VBA évoqué en [section 20.1](01-modele-securite-access.md), trivial à contourner.

Un fichier `.accde` est une version compilée et exécutable seule de la base `.accdb`. Lors de sa génération, Access supprime le code source VBA et n'en conserve que la forme compilée, directement exécutable mais illisible. C'est l'équivalent moderne, pour le format `.accdb`, de l'ancien `.mde` qui s'appliquait aux fichiers `.mdb`. Il ne faut pas le confondre avec le format `.accdr`, qui ne verrouille rien mais force simplement l'ouverture en mode runtime (voir [chapitre 21](../21-deploiement-distribution/02-access-runtime.md)).

## Ce que protège — et ce que ne protège pas — un fichier ACCDE

Comprendre précisément le périmètre du `.accde` est indispensable pour ne pas se reposer sur une protection qu'il n'offre pas.

Du côté de ce qui est verrouillé : le code VBA n'est plus consultable ni modifiable, puisque le source a disparu. De même, les formulaires et les états ne peuvent plus être ouverts en mode Création ; on ne peut donc plus modifier leur conception, ni en créer, ni en supprimer. Enfin, l'import ou l'export de formulaires, d'états et de modules vers ou depuis une autre base est bloqué, ce qui empêche de récupérer ces objets ailleurs.

Du côté de ce qui reste accessible, en revanche, il faut être lucide. Les tables, les requêtes, les macros et les relations demeurent visibles et modifiables : le `.accde` ne fige pas la structure des données ni la logique des requêtes sauvegardées. Surtout, les données elles-mêmes restent entièrement accessibles, car le `.accde` ne chiffre rien. La protection des données reste donc le rôle exclusif du mot de passe et du chiffrement vus en section 20.2.

En résumé, le `.accde` protège le code et la conception de l'interface, mais ni les données ni la structure des tables et des requêtes. C'est pourquoi il se combine presque toujours avec d'autres mesures.

## Les prérequis avant compilation

La génération d'un `.accde` n'aboutit que si plusieurs conditions sont réunies. La base doit d'abord être au format actuel `.accdb` ; un ancien fichier `.mdb` devra être converti au préalable. Ensuite, et c'est la cause d'échec la plus fréquente, la totalité du code VBA doit compiler sans erreur : la moindre erreur de compilation interrompt l'opération. Il est donc recommandé de lancer au préalable une compilation explicite depuis l'éditeur VBA, par le menu Débogage puis Compiler, afin de détecter et corriger les éventuelles erreurs. Enfin, toutes les références de bibliothèques doivent être valides ; une référence cassée (signalée par la mention MANQUANTE dans la liste des références) empêche elle aussi la compilation.

## Créer un ACCDE via l'interface

Une fois ces prérequis satisfaits, la création se fait depuis l'interface. On se rend dans l'onglet Fichier, puis dans Enregistrer sous, on choisit Enregistrer la base de données sous et l'on sélectionne l'option Créer ACCDE. Access demande alors l'emplacement et le nom du fichier de sortie, puis génère un nouveau fichier portant l'extension `.accde`. Le fichier `.accdb` d'origine, lui, est conservé intact.

## La règle d'or : conserver précieusement le fichier source ACCDB

Voici le point le plus important de toute cette section, à ne jamais perdre de vue : la compilation en `.accde` est une opération à sens unique. Il n'existe aucun moyen, ni par l'interface ni par un quelconque outil officiel, de reconvertir un `.accde` en `.accdb` pour en récupérer le code source. Le source n'est pas masqué, il est supprimé.

La conséquence pratique est capitale. Le fichier `.accdb` d'origine devient votre unique source maîtresse, et toute évolution future de l'application doit impérativement être réalisée sur ce `.accdb`, après quoi vous régénérez le `.accde` à distribuer. Perdre le `.accdb` revient à perdre définitivement la capacité de modifier l'application. Cela impose une discipline de sauvegarde rigoureuse du fichier source, idéalement couplée à un versioning, sujet traité en [section 24.4](../24-bonnes-pratiques-ressources/04-versioning-git.md). Ne distribuez jamais le `.accdb` : il reste chez le développeur ; seul le `.accde` est livré.

## Compatibilité : version et architecture 32 / 64 bits

Un fichier `.accde` n'est pas librement interchangeable entre toutes les configurations. Il est prudent de le générer dans la version d'Access la plus ancienne que vous devez prendre en charge, car un `.accde` créé dans une version récente peut refuser de s'ouvrir dans une version antérieure. De même, l'architecture compte : un `.accde` compilé en 32 bits et un `.accde` compilé en 64 bits ne sont pas équivalents, et un fichier prévu pour une architecture peut ne pas fonctionner sur l'autre, en particulier lorsque le code contient des déclarations d'API Windows. Veillez donc à produire un `.accde` correspondant à la version et à l'architecture de l'environnement cible. Ces questions de compatibilité sont approfondies aux sections [21.6](../21-deploiement-distribution/06-compatibilite-versions-access.md) et [21.7](../21-deploiement-distribution/07-32-bits-vs-64-bits.md).

## Combiner l'ACCDE avec les autres protections

Le `.accde` prend toute sa valeur lorsqu'il est intégré à une stratégie globale. La combinaison la plus courante associe le `.accde`, qui protège le code, au chiffrement par mot de passe, qui protège les données : les deux sont parfaitement cumulables sur un même fichier et se complètent idéalement pour une distribution.

Dans une architecture séparée front-end / back-end, la répartition naturelle est la suivante : le front-end, qui concentre le code, les formulaires et les états, est distribué sous forme de `.accde` ; le back-end, qui ne contient que des données et peu ou pas de code, n'a pas besoin d'être compilé mais sera protégé par mot de passe.

On renforce généralement encore le verrouillage par des réglages de démarrage qui masquent le volet de navigation, désactivent les touches spéciales et restreignent l'interface, comme décrit en [section 17.7](../17-interface-utilisateur-avancee/07-options-demarrage.md). Ces options empêchent un utilisateur d'accéder directement aux tables ou de contourner l'interface, et complètent utilement la protection du code apportée par le `.accde`.

Il peut être pratique de faire en sorte que l'application détecte elle-même si elle s'exécute sur la version compilée, par exemple pour masquer des menus réservés au développement. Le moyen le plus simple consiste à inspecter l'extension du fichier courant :

```vba
Public Function EstVersionCompilee() As Boolean
    ' True si l'application tourne sur un fichier .accde
    EstVersionCompilee = (LCase$(Right$(CurrentProject.Name, 6)) = ".accde")
End Function
```

## Et par code ?

À la différence du mot de passe, la génération d'un `.accde` ne dispose pas de méthode VBA simple et officielle : c'est avant tout une opération interactive, réalisée depuis l'interface d'Access. Son automatisation fiable, dans le cadre d'un processus de fabrication ou de déploiement, n'est pas triviale et relève généralement de l'outillage de build ou du pilotage d'Access en automation, abordé sous l'angle du déploiement en [section 21.8](../21-deploiement-distribution/08-installation-automatisee.md). Pour la grande majorité des projets, produire le `.accde` à la main au moment de la livraison reste l'approche normale.

## Limites et bonnes pratiques

Le `.accde` constitue une protection sérieuse et recommandée du code : le source étant réellement retiré, il n'existe pas de voie officielle pour le récupérer. Il convient toutefois de rester mesuré et de garder en tête ses limites. Il ne protège ni les données ni la structure des tables, d'où la nécessité de le combiner au chiffrement. Et puisqu'il devient impossible de modifier l'application livrée, il est impératif de tester soigneusement le `.accde` avant distribution, dans des conditions proches de la production : un défaut découvert après livraison ne se corrige qu'en repartant du `.accdb` source pour régénérer un nouveau fichier.

## Ce qu'il faut retenir

La compilation en `.accde` est la bonne réponse à la question de la protection du code VBA et de la conception des formulaires et états. Elle supprime le source plutôt que de le masquer, ce qui la rend bien plus solide que le mot de passe de projet VBA. Sa contrepartie est son irréversibilité : le fichier `.accdb` source doit être conservé et versionné avec le plus grand soin, car il devient la seule base modifiable. Le `.accde` ne protégeant pas les données, on l'associe systématiquement au chiffrement par mot de passe et, le plus souvent, à un verrouillage de l'interface au démarrage. Ainsi combiné, il forme la couche « protection du code » d'une stratégie de sécurité cohérente.

---


⏭️ [20.4. Signature numérique des macros et confiance du code](/20-securite-protection/04-signature-numerique.md)
