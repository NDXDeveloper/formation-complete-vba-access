🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 11. SQL dans Access VBA

Après avoir étudié les deux grandes technologies d'accès aux données — DAO (chapitre 9) et ADO (chapitre 10) —, ce chapitre se penche sur le **langage** que ces technologies exécutent : le **SQL**. Car derrière chaque recordset ouvert, chaque requête lancée, chaque mise à jour de masse, c'est toujours du SQL qui travaille. Maîtriser le SQL depuis VBA, c'est gagner en puissance et en souplesse : construire des requêtes à la volée, traiter des milliers d'enregistrements en une instruction, automatiser la création de tables, interroger des serveurs distants.

Ce chapitre couvre le SQL tel qu'on l'emploie réellement dans une application Access pilotée par VBA : comment l'exécuter, comment le construire dynamiquement sans danger, et quelles sont les spécificités du dialecte propre à Access.

---

## SQL : le langage caché d'Access

Beaucoup d'utilisateurs d'Access créent des requêtes sans jamais écrire une ligne de SQL, à l'aide du concepteur visuel (la grille QBE). Pourtant, **toute requête Access est du SQL** sous le capot : le concepteur graphique ne fait que produire une instruction SQL, consultable à tout moment en basculant en *mode SQL*. Comprendre ce langage, c'est donc accéder à la mécanique réelle d'Access — et dépasser les limites de l'interface graphique.

Access n'emploie pas un SQL totalement standard, mais un dialecte appelé **Jet SQL** (ou ACE SQL selon le moteur). Largement compatible avec le standard, il possède ses propres fonctions, ses particularités de syntaxe et ses limites — autant d'éléments que ce chapitre détaille, car ils réservent des surprises à qui vient d'un autre SGBD.

---

## Un chapitre transversal à DAO et ADO

Les chapitres 9 et 10 décrivaient les **objets** d'accès aux données ; celui-ci décrit le **langage** qu'ils exécutent. C'est ce qui rend le chapitre 11 transversal : le SQL est le fil commun à toutes les approches. Une même instruction peut être lancée de plusieurs façons, selon la technologie retenue.

```vba
' Une même instruction SQL, trois façons de l'exécuter

' 1. Via DoCmd (objet d'automatisation d'Access)
DoCmd.RunSQL "UPDATE Clients SET Actif = True WHERE Region = 'Nord';"

' 2. Via DAO (chapitre 9)
CurrentDb.Execute "UPDATE Clients SET Actif = True WHERE Region = 'Nord';", dbFailOnError

' 3. Via ADO (chapitre 10)
CurrentProject.Connection.Execute "UPDATE Clients SET Actif = True WHERE Region = 'Nord';"
```

Ces trois lignes produisent le même effet. Choisir la bonne méthode — et comprendre leurs différences de comportement — est précisément l'objet de la première section.

---

## Deux compétences distinctes

Travailler le SQL depuis VBA mobilise en réalité **deux savoir-faire complémentaires**, qu'il convient de ne pas confondre.

Le premier est l'**écriture d'un SQL correct** : connaître la syntaxe, les fonctions et les limites du dialecte Jet, savoir formuler une jointure, une sous-requête, une requête d'union. C'est la maîtrise du *langage*.

Le second est l'**exécution propre depuis VBA** : choisir la méthode d'exécution adaptée, paramétrer les requêtes pour les sécuriser, construire dynamiquement une instruction sans introduire de faille ni d'erreur de formatage. C'est la maîtrise de l'*outillage*.

Ce chapitre développe les deux, car l'un sans l'autre mène à des applications soit fragiles, soit limitées.

---

## Pourquoi exécuter du SQL depuis VBA ?

Plusieurs besoins justifient de recourir au SQL plutôt qu'à la seule manipulation par recordset.

Les **requêtes dynamiques** d'abord : construire une instruction adaptée à la saisie de l'utilisateur, à des critères variables, à un contexte d'exécution. Un filtre choisi à l'écran, une période sélectionnée, une liste de critères combinés se traduisent naturellement en SQL bâti à la volée.

Les **opérations de masse** ensuite : comme nous l'avons souligné à la section 10.5, une requête action (`UPDATE`, `DELETE`, `INSERT`) traitant des milliers de lignes est radicalement plus rapide qu'une boucle sur un recordset. Le SQL ensembliste est l'outil de la performance.

L'**automatisation et la gestion de la structure** enfin : créer ou modifier des tables, des index, des contraintes par code (DDL) permet de faire évoluer une base sans intervention manuelle. Le tout dans un **langage commun** qui fonctionne indifféremment avec DAO et ADO.

---

## Ce que vous allez apprendre dans ce chapitre

Le chapitre progresse de l'exécution la plus simple jusqu'aux constructions les plus avancées.

La section **11.1** compare les deux grandes voies d'exécution du SQL depuis VBA, **`DoCmd.RunSQL`** et **`CurrentDb.Execute`**, et explique comment choisir entre elles. La section **11.2** traite des requêtes **`SELECT`** et de l'ouverture d'un recordset sur leur résultat. La section **11.3** couvre les **requêtes action** — `INSERT`, `UPDATE`, `DELETE` — qui modifient les données.

La section **11.4** aborde le **DDL** (`CREATE TABLE`, `ALTER TABLE`, `DROP TABLE`) pour modifier la structure de la base par code. La section **11.5** est consacrée au **SQL paramétré** et à la prévention des **injections SQL**, un impératif de sécurité. La section **11.6** détaille la **construction dynamique** de SQL, ses bonnes pratiques et ses pièges.

La section **11.7** s'attaque à un écueil notoire d'Access : la **localisation et le formatage** dans le SQL dynamique — dates au format US/ISO, séparateur décimal, échappement des apostrophes. La section **11.8** explore le **dialecte Jet SQL**, ses spécificités et ses limites face au SQL standard. La section **11.9** présente les **requêtes Pass-Through** qui envoient du SQL natif directement à SQL Server, prolongeant le chapitre 10.9.

La section **11.10** couvre les **fonctions de domaine** (`DLookup`, `DSum`, `DCount`…), pratiques pour interroger ponctuellement les données. Enfin, la section **11.11** traite des **requêtes d'union, des sous-requêtes et des expressions complexes**, pour les besoins les plus élaborés.

---

## Le fil rouge : sécurité et localisation

Deux thèmes traversent l'ensemble du chapitre et méritent une vigilance particulière, car ils sont à l'origine de la majorité des bugs et failles rencontrés en pratique.

Le premier est la **sécurité**. Construire une requête en concaténant des valeurs saisies par l'utilisateur ouvre la porte aux **injections SQL** — une faille aussi répandue que dangereuse. Les sections 11.5 et 11.6 montrent comment s'en prémunir, en privilégiant le paramétrage.

Le second est la **localisation**. Access piège régulièrement les développeurs sur le **formatage des dates et des nombres** dans le SQL dynamique : une date doit y figurer au format américain ou ISO et entre dièses (`#`), le séparateur décimal doit être le point, et les apostrophes doivent être doublées. La section 11.7 y est entièrement dédiée. Ignorer ces règles produit des erreurs subtiles, qui ne se manifestent parfois que sur certains postes selon leurs paramètres régionaux.

Gardez ces deux fils rouges à l'esprit dès maintenant : ils conditionnent la robustesse de tout code SQL dynamique.

---

## Prérequis

Ce chapitre suppose acquis les fondamentaux de VBA (chapitre 3), et tout particulièrement la **manipulation des chaînes de caractères** (chapitre 3.6), omniprésente dans la construction dynamique de SQL. Une connaissance des technologies d'accès aux données — DAO (chapitre 9) et ADO (chapitre 10) — est également nécessaire, puisque c'est par elles que le SQL est exécuté. Une familiarité de base avec le SQL est utile, mais le chapitre prend soin d'expliquer les aspects propres à Access.

---

## Note de cadrage et ressources de référence

Ce chapitre se concentre sur le SQL **tel qu'on l'emploie depuis VBA dans Access** : son exécution, sa construction dynamique, son formatage, et les spécificités du dialecte Jet. Il ne constitue pas un cours de SQL exhaustif partant de zéro, mais il couvre l'essentiel des aspects qu'un développeur Access doit maîtriser au quotidien.

Pour une consultation de référence, deux annexes complètent ce chapitre : l'**annexe E** détaille la syntaxe et les fonctions du **Jet SQL**, tandis que l'**annexe F** établit la **correspondance entre Jet SQL et T-SQL** (SQL Server) — précieuse lors d'une migration ou d'un travail en architecture hybride (chapitre 23).

---

La première étape, avant d'écrire la moindre requête sophistiquée, consiste à savoir l'exécuter correctement. Or les deux méthodes principales — `DoCmd.RunSQL` et `CurrentDb.Execute` — ne se valent pas : leurs différences de comportement, notamment en matière de messages d'avertissement et de détection d'erreurs, ont des conséquences concrètes. C'est par là que commence le chapitre.


⏭️ [11.1. Exécution de SQL par DoCmd.RunSQL et CurrentDb.Execute](/11-sql-access-vba/01-execution-sql-runsql-execute.md)
