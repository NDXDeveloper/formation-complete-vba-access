🔝 Retour au [Sommaire](/SOMMAIRE.md)

# J. Glossaire des termes Access et base de données

Ce glossaire définit les principaux termes employés tout au long de la formation. Il sert à lever toute ambiguïté de vocabulaire. Les définitions sont volontairement brèves ; chaque notion est développée dans le chapitre correspondant.

> ℹ️ Le terme anglais ou son synonyme est indiqué entre parenthèses lorsqu'il aide au repérage, l'écosystème Access mêlant vocabulaire français et anglais.

## A

- **ACCDB** — Format de fichier de base de données Access, en vigueur depuis la version 2007.
- **ACCDE** — Version compilée et verrouillée d'une base `.accdb`, dont le code VBA n'est plus modifiable.
- **ACCDR** — Base `.accdb` configurée pour s'ouvrir en mode runtime (exécution sans modification).
- **ACE (Access Database Engine)** — Moteur de base de données intégré à Access depuis 2007, successeur de Jet, gérant les fichiers `.accdb`.
- **ADO (ActiveX Data Objects)** — Bibliothèque d'accès aux données via OLE DB, alternative à DAO, particulièrement adaptée aux sources externes.
- **ADP (Access Data Project)** — Ancien format de projet connecté directement à SQL Server, aujourd'hui abandonné.
- **Argument** — Valeur transmise à une procédure, une fonction ou une méthode lors de son appel.
- **Atomicité** — Propriété d'une transaction qui s'exécute entièrement ou pas du tout.

## B

- **Back-end (dorsal)** — Dans une architecture scindée, le fichier contenant les tables (les données), partagé sur le réseau.
- **Binding (liaison)** — Mécanisme reliant une variable objet à son type COM ; précoce à la compilation ou tardif à l'exécution.
- **BOF / EOF** — *Begin/End Of File* : marqueurs indiquant qu'un Recordset est positionné avant le premier ou après le dernier enregistrement.
- **Bookmark (signet)** — Référence identifiant de façon unique la position d'un enregistrement dans un Recordset.
- **Bound (lié)** — Se dit d'un contrôle ou d'un formulaire rattaché à une source de données.

## C

- **Champ (Field)** — Colonne d'une table, représentant un attribut des enregistrements.
- **Classe (module de)** — Modèle définissant les propriétés et méthodes d'objets que l'on instancie.
- **Clé étrangère (Foreign Key)** — Champ référençant la clé primaire d'une autre table, support des relations.
- **Clé primaire (Primary Key)** — Champ ou ensemble de champs identifiant de façon unique chaque enregistrement.
- **Collection** — Objet regroupant un ensemble d'objets de même nature (ex. `Forms`, `Controls`).
- **Compactage** — Opération de maintenance réduisant la taille du fichier et réorganisant les données.
- **Connection (ADO)** — Objet représentant une session ouverte avec une source de données.
- **Contrôle (Control)** — Élément d'interface d'un formulaire ou d'un état (zone de texte, bouton, liste…).
- **CurrentDb** — Fonction DAO renvoyant une référence à la base de données ouverte.
- **Curseur (Cursor)** — En ADO, mécanisme de parcours d'un Recordset déterminant la navigation et les possibilités de mise à jour.

## D

- **DAL (Data Access Layer)** — Couche logicielle isolant l'accès aux données du reste de l'application.
- **DAO (Data Access Objects)** — Bibliothèque native d'accès aux données du moteur ACE/Jet.
- **DBEngine** — Objet racine de la hiérarchie d'objets DAO.
- **DDL (Data Definition Language)** — Instructions SQL définissant la structure (`CREATE`, `ALTER`, `DROP`).
- **Dirty** — État d'un enregistrement ou d'un contrôle modifié mais non encore enregistré.
- **DoCmd** — Objet exposant les actions d'automatisation d'Access, équivalent des actions de macro.
- **Domaine (fonctions de)** — `DLookup`, `DSum`, `DCount`… agrégeant des valeurs sur un ensemble d'enregistrements sans requête explicite.
- **Dynaset** — Type de Recordset modifiable reflétant les changements de la base (jeu dynamique).

## E

- **Early Binding (liaison précoce)** — Liaison d'un objet à son type dès la compilation, via une référence COM ; active l'IntelliSense.
- **Enregistrement (Record)** — Ligne d'une table, regroupant les valeurs des champs pour une entité.
- **Énumération (Enum)** — Ensemble nommé de constantes entières liées.
- **État (Report)** — Objet de présentation et d'impression des données.
- **Événement (Event)** — Action déclenchant l'exécution d'une procédure (`Click`, `Open`, `Current`…).

## F

- **Feuille de données (Datasheet)** — Affichage tabulaire des données, en lignes et colonnes.
- **Formulaire (Form)** — Objet d'interface destiné à la saisie, à la consultation et à la navigation.
- **Fonction (Function)** — Procédure nommée renvoyant une valeur.
- **Front-end (frontal)** — Dans une architecture scindée, le fichier contenant l'interface (formulaires, états, code), copié sur chaque poste.

## I

- **Index** — Structure accélérant les recherches et les tris sur un ou plusieurs champs.
- **Instance / Instanciation** — Objet concret créé à partir d'une classe ; action de le créer.
- **Intégrité référentielle** — Ensemble de règles garantissant la cohérence des relations entre tables.
- **IntelliSense** — Aide à la saisie de l'éditeur (listes de membres, info-bulles, complétion).

## J

- **Jet** — Ancien moteur de base de données d'Access (jusqu'à 2003), remplacé par ACE.
- **Jet SQL** — Dialecte SQL du moteur Jet/ACE, propre à Access.
- **Jointure (Join)** — Opération combinant les lignes de plusieurs tables selon une condition.

## L

- **Late Binding (liaison tardive)** — Liaison d'un objet à son type à l'exécution via `CreateObject` ; sans IntelliSense, mais souple au déploiement.
- **LDB / LACCDB** — Fichier de verrouillage créé à l'ouverture d'une base partagée, listant les sessions actives.
- **LongPtr** — Type entier de la taille d'un pointeur (4 octets en 32 bits, 8 en 64 bits), pour les handles d'API.

## M

- **Macro (Access)** — Suite d'actions définies sans code, distincte du VBA.
- **MDB** — Ancien format de fichier Access associé au moteur Jet.
- **Méthode (Method)** — Action qu'un objet sait effectuer.
- **Module** — Conteneur de code VBA (standard, de classe, ou associé à un formulaire/état).
- **Multivalué (champ — MVF)** — Champ stockant plusieurs valeurs par enregistrement.

## N

- **Normalisation** — Organisation des tables visant à réduire la redondance et à garantir la cohérence des données.
- **Null** — Absence de valeur, distincte de zéro ou de la chaîne vide.
- **NuméroAuto (AutoNumber)** — Champ générant automatiquement une valeur unique, généralement croissante.

## O

- **Objet (Object)** — Entité dotée de propriétés et de méthodes.
- **ODBC (Open Database Connectivity)** — Interface standard d'accès aux SGBD via des pilotes.
- **OLE DB** — Interface d'accès aux données de Microsoft, sous-jacente à ADO.
- **OpenArgs** — Argument transmis à un formulaire ou à un état au moment de son ouverture.
- **Optimiste (verrouillage)** — Stratégie ne verrouillant l'enregistrement qu'au moment de l'écriture.

## P

- **Paramètre (Parameter)** — Valeur transmise à une procédure ou à une requête.
- **Pass-Through (requête)** — Requête envoyée telle quelle à un serveur (ex. SQL Server) pour exécution côté serveur.
- **Pessimiste (verrouillage)** — Stratégie verrouillant l'enregistrement dès le début de l'édition.
- **Pilote (Driver)** — Composant logiciel assurant la communication avec un SGBD (ODBC).
- **POO** — Programmation Orientée Objet : paradigme fondé sur les classes et les objets.
- **Portée (Scope)** — Domaine de visibilité d'une variable (locale, module, globale).
- **Procédure (Sub)** — Bloc d'instructions nommé ne renvoyant pas de valeur.
- **Propriété (Property)** — Caractéristique d'un objet, lisible et parfois modifiable.
- **PtrSafe** — Mot-clé marquant une déclaration d'API compatible 64 bits.

## Q

- **QueryDef** — Objet DAO représentant une requête enregistrée.
- **Requête (Query)** — Interrogation ou action sur les données, exprimée en SQL.
- **Requête action** — Requête modifiant les données (`INSERT`, `UPDATE`, `DELETE`) ou la structure.

## R

- **RecordSource** — Propriété définissant la source de données d'un formulaire ou d'un état.
- **Recordset** — Ensemble d'enregistrements manipulable par code (DAO ou ADO).
- **RecordsetClone** — Copie du Recordset d'un formulaire, permettant de parcourir les enregistrements sans déplacer l'affichage.
- **Référence (Reference)** — Lien vers une bibliothèque COM activée dans le projet VBA.
- **Relation** — Association définie entre deux tables via leurs clés.
- **Rollback (annulation)** — Retour à l'état antérieur d'une transaction non validée.
- **Runtime (Access Runtime)** — Version gratuite et restreinte d'Access, destinée à exécuter une application sans la modifier.

## S

- **SGBD** — Système de Gestion de Base de Données.
- **Snapshot (capture instantanée)** — Type de Recordset en lecture seule, figé à l'ouverture.
- **Sous-formulaire (Subform)** — Formulaire imbriqué dans un autre, souvent en relation parent/enfant.
- **SQL (Structured Query Language)** — Langage d'interrogation et de manipulation des données.

## T

- **Table** — Structure de stockage des données, organisée en champs et enregistrements.
- **TableDef** — Objet DAO décrivant la structure d'une table.
- **TempVars** — Variables globales de session propres à Access.
- **Transaction** — Ensemble d'opérations exécutées comme un tout indivisible.
- **T-SQL (Transact-SQL)** — Dialecte SQL de Microsoft SQL Server.

## U

- **Unbound (indépendant)** — Contrôle ou formulaire non rattaché à une source de données.
- **UNC (chemin)** — Notation réseau `\\serveur\partage\…` désignant une ressource partagée.

## V

- **Variable** — Emplacement nommé stockant une valeur typée.
- **Variant** — Type VBA pouvant contenir n'importe quel type de donnée.
- **Verrouillage (Locking)** — Mécanisme empêchant les accès concurrents conflictuels.

## W

- **Workspace** — Objet DAO représentant une session de travail, support des transactions.

---

> 💡 **À retenir.** Quelques distinctions reviennent souvent et méritent d'être claires : **front-end / back-end** (interface vs données), **DAO / ADO** (accès natif vs externe), **lié / indépendant** (bound vs unbound), **liaison précoce / tardive** (early vs late binding), **Dynaset / Snapshot** (modifiable vs lecture seule) et **Jet SQL / T-SQL** (dialecte Access vs SQL Server). Chacune est approfondie dans le chapitre qui s'y rapporte.

⏭️ [K. Modèles de code réutilisables (snippets)](/annexes/k-modeles-code-reutilisables.md)
