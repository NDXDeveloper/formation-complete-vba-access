🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 10.1. DAO vs ADO — quelle technologie choisir ?

C'est sans doute la question la plus fréquente — et la plus mal documentée — que se pose un développeur Access lorsqu'il aborde l'accès aux données par code. Une simple recherche en ligne renvoie des avis contradictoires : certains affirment que DAO est obsolète, d'autres qu'ADO est dépassé, et chacun semble avoir raison… selon l'époque où l'article a été écrit. Cette section a pour but de dissiper la confusion et de vous fournir des critères de décision clairs, fondés sur la réalité technique d'Access aujourd'hui.

La bonne nouvelle, annoncée dès l'introduction du chapitre, c'est qu'il n'existe pas de « gagnant » universel : DAO et ADO répondent à des besoins différents, et un développeur Access aguerri sait passer de l'un à l'autre selon le contexte.

---

## Deux philosophies opposées

Pour comprendre quand utiliser chaque technologie, il faut d'abord saisir l'intention derrière chacune.

**DAO** (*Data Access Objects*) a été conçu spécifiquement pour le moteur de base de données d'Access — d'abord Jet, puis ACE. C'est une technologie « maison », taillée sur mesure pour exploiter au plus près les capacités du moteur natif. DAO connaît intimement Access : ses champs multivalués, ses pièces jointes, ses index, ses relations. En contrepartie, son horizon s'arrête largement aux frontières du moteur ACE.

**ADO** (*ActiveX Data Objects*) procède d'une ambition inverse. Né de la stratégie « *Universal Data Access* » de Microsoft, il vise à offrir une interface unique et homogène vers *n'importe quelle* source de données, grâce à la couche d'abstraction OLE DB. ADO ne connaît pas Access en particulier ; il connaît une multitude de fournisseurs de données, parmi lesquels celui d'Access n'est qu'un cas parmi d'autres. Sa force est l'ouverture ; sa faiblesse, une moindre intimité avec les spécificités du moteur ACE.

En une phrase : **DAO est un spécialiste d'Access, ADO est un généraliste des bases de données.**

---

## Un peu d'histoire pour comprendre la cacophonie

Si la documentation en ligne se contredit autant, c'est parce que la position officielle de Microsoft a elle-même changé plusieurs fois.

À l'origine (Access 1.0, au début des années 1990), **DAO était la seule technologie** d'accès aux données. Puis, à la fin des années 1990, Microsoft lance ADO dans le cadre de sa vision d'un accès universel aux données. Avec Access 2000, l'éditeur **pousse activement ADO** : les nouvelles bases référencent ADO par défaut, et la documentation présente DAO comme une technologie héritée, vouée à disparaître. Beaucoup d'articles datant de cette période recommandent donc d'abandonner DAO — et certains circulent encore aujourd'hui.

Le vent a tourné avec **Access 2007** et l'arrivée du moteur **ACE** (format `.accdb`). Microsoft a alors **rétabli DAO comme la technologie native et recommandée** pour travailler avec le moteur d'Access, en livrant une bibliothèque DAO mise à jour pour ACE. Dans le même mouvement, les *Access Data Projects* (ADP) — une architecture entièrement bâtie sur ADO et liée à SQL Server — ont été progressivement abandonnés, puis purement supprimés dans Access 2013. Quant à *ODBCDirect*, une fonctionnalité de DAO qui permettait d'attaquer directement des sources ODBC, elle a disparu dès Access 2007.

La morale de cette histoire : **les recommandations « DAO est mort » appartiennent au passé**. Aujourd'hui, DAO est pleinement vivant et constitue le choix par défaut pour le moteur ACE, tandis qu'ADO conserve un rôle bien défini pour l'ouverture vers l'extérieur. Méfiez-vous donc de la date des articles que vous consultez.

---

## Comparaison technique, critère par critère

### Modèle objet

DAO expose une **hiérarchie profonde et riche** (`DBEngine → Workspace → Database → Recordset`, avec aussi `TableDef`, `QueryDef`, `Relation`, `Index`…). Cette richesse reflète tout ce que le moteur ACE sait faire. ADO, à l'inverse, propose un **modèle plat et compact** articulé autour de trois objets (`Connection`, `Command`, `Recordset`). Le modèle ADO est plus rapide à apprendre, mais moins expressif sur les particularités d'Access.

### Performance sur le moteur ACE local

Pour des opérations sur une base locale ou liée au moteur ACE, **DAO est généralement plus rapide**. La raison est structurelle : DAO dialogue directement avec le moteur, alors qu'ADO ajoute la couche intermédiaire OLE DB. Pour un parcours intensif de recordsets locaux, cette surcouche se paie en performances. Sur une base purement Access, le réflexe DAO est donc justifié.

### Fonctionnalités propres à Access

C'est ici que l'écart se creuse nettement en faveur de DAO. Les **champs multivalués** et les **champs Pièce jointe** (via `Recordset2` et `Field2`, voir chapitre 9.13) ne sont **pas pris en charge nativement par ADO**. De même, la manipulation fine de la structure de la base — créer des `TableDef`, des `QueryDef`, gérer les `Relations` (chapitre 12) — relève naturellement de DAO. ADO peut faire du DDL, mais via une bibliothèque distincte, **ADOX** (*ActiveX Data Objects Extensions*), au prix d'une complexité supplémentaire et d'une couverture moins complète.

### Accès aux sources de données externes

Le rapport de force s'inverse complètement. Pour se connecter à **SQL Server, Oracle, PostgreSQL, MySQL** et consorts, **ADO est l'outil de référence** : il suffit de changer de fournisseur OLE DB dans la chaîne de connexion. DAO, lui, n'accède aux sources externes qu'*indirectement*, via des tables liées gérées par le moteur ACE ou via des requêtes pass-through. Dès qu'une application doit attaquer directement un SGBD distant sans table liée, ADO s'impose.

### Recordsets déconnectés et persistance

ADO offre une capacité qu'**DAO ne possède tout simplement pas** : le *recordset déconnecté*. En positionnant l'emplacement du curseur côté client (`adUseClient`), on peut charger des données, fermer la connexion, puis continuer à manipuler le jeu d'enregistrements en mémoire — et même le **sauvegarder dans un fichier XML** pour le recharger plus tard (voir section 10.7). DAO n'a pas d'équivalent : un recordset DAO reste lié à sa base.

### Liaison aux formulaires et aux états

Dans une base `.accdb` standard, la couche de données des formulaires et des états repose sur le moteur ACE et **expose des recordsets DAO**. En particulier, la propriété `RecordsetClone` d'un formulaire (chapitre 9.12) renvoie un `DAO.Recordset`. Pour manipuler par code les enregistrements affichés dans un formulaire, **DAO est donc le partenaire naturel**. La liaison d'un recordset ADO à un formulaire reste possible dans des cas limités, mais c'était surtout une caractéristique des ADP aujourd'hui disparus ; ce n'est pas le scénario standard d'une application `.accdb`.

### Références et configuration

Côté DAO, on active la référence **« Microsoft Office xx.x Access Database Engine Object Library »** et l'on préfixe les objets par `DAO.`. Côté ADO, on active **« Microsoft ActiveX Data Objects 6.1 Library »** et l'on préfixe par `ADODB.`. Ces préfixes ne sont pas une coquetterie : ils sont indispensables pour lever l'ambiguïté entre objets homonymes — `Recordset` existant dans les deux bibliothèques (voir chapitre 2.5 sur les références).

---

## Tableau de synthèse

| Critère | DAO | ADO |
|---|---|---|
| **Conçu pour** | Le moteur Jet/ACE d'Access | Toute source via OLE DB |
| **Modèle objet** | Hiérarchique et riche | Plat et compact |
| **Performance sur ACE local** | Optimale (accès direct) | Bonne, mais surcouche OLE DB |
| **Sources externes (SQL Server, Oracle…)** | Indirect (tables liées, pass-through) | Natif et étendu |
| **Champs multivalués / pièces jointes** | Pris en charge (`Recordset2`/`Field2`) | Non pris en charge nativement |
| **Structure (TableDef, QueryDef, Relations)** | Complet et direct | Partiel, via ADOX |
| **`RecordsetClone` des formulaires** | Oui (c'est un `DAO.Recordset`) | Non (formulaires `.accdb` = DAO) |
| **Recordsets déconnectés** | Non | Oui (`adUseClient`) |
| **Persistance XML d'un recordset** | Non | Oui (méthode `Save`) |
| **Recherche indexée (`Seek`)** | Oui, directe (recordset type Table) | Possible, mais conditionnée |
| **Référence à activer** | Microsoft … Access Database Engine Object Library | Microsoft ActiveX Data Objects 6.1 Library |
| **Préfixe recommandé** | `DAO.` | `ADODB.` |
| **Statut actuel** | Native et recommandée pour ACE | Stable, dédiée à l'ouverture externe |

---

## Critères de décision

Plutôt qu'une règle unique, raisonnez par scénario. Les situations suivantes orientent assez clairement le choix.

**Choisissez DAO** lorsque votre application travaille principalement sur une base locale ou liée au moteur ACE — ce qui correspond à la grande majorité des applications Access. DAO s'impose également dès que vous manipulez des champs multivalués ou des pièces jointes, que vous intervenez sur la structure de la base (tables, requêtes, relations), ou que vous travaillez sur les enregistrements d'un formulaire via `RecordsetClone`. En résumé : **par défaut, dans un univers Access, c'est DAO.**

**Choisissez ADO** lorsque vous devez vous connecter directement à un SGBD externe sans passer par des tables liées, lorsque vous avez besoin de recordsets déconnectés ou de persistance XML, ou lorsque votre application s'inscrit dans une architecture distribuée ou une trajectoire de migration vers SQL Server (chapitre 23). ADO s'impose aussi, plus prosaïquement, lorsque vous **maintenez du code existant déjà écrit en ADO** : il n'y a aucun bénéfice à le réécrire en DAO sans raison.

Entre ces deux pôles, le bon réflexe est de partir de la question : *« mes données vivent-elles dans le moteur ACE, ou ailleurs ? »* Si la réponse est « dans ACE », DAO est presque toujours le meilleur point de départ. Si c'est « ailleurs » — ou « bientôt ailleurs » — ADO mérite considération.

---

## Peut-on utiliser les deux ? Oui, et c'est fréquent

Une idée reçue voudrait qu'il faille choisir son camp une fois pour toutes. Il n'en est rien : **DAO et ADO cohabitent sans problème** dans une même application, et même dans un même module, voire une même procédure. Une application réaliste utilise souvent DAO pour l'essentiel du travail sur les données locales, et bascule ponctuellement sur ADO pour un échange avec un serveur distant.

La seule précaution impérative est la **qualification explicite des types**, car certains noms d'objets sont communs aux deux bibliothèques. Sans préfixe, l'objet retenu dépend de l'ordre des références dans le projet — source de bugs déroutants.

```vba
' Toujours préfixer pour éviter toute ambiguïté
Dim rsLocal As DAO.Recordset      ' recordset DAO, sans équivoque
Dim rsDistant As ADODB.Recordset  ' recordset ADO, sans équivoque

' À PROSCRIRE : le type effectif dépend de l'ordre des références
Dim rsAmbigu As Recordset          ' DAO ou ADO ? Imprévisible
```

Cette discipline de nommage, déjà évoquée au chapitre 2.5, n'est pas négociable dès lors que les deux bibliothèques sont référencées simultanément.

---

## En résumé

Retenez trois idées de cette section. Premièrement, **DAO n'est pas obsolète** : c'est la technologie native et recommandée pour le moteur ACE, et le choix par défaut d'une application Access classique. Deuxièmement, **ADO n'est pas dépassé non plus** : il reste irremplaçable pour l'ouverture vers des sources externes, les recordsets déconnectés et les architectures distribuées. Troisièmement, **les deux peuvent — et doivent parfois — coexister**, à condition de qualifier explicitement les types.

Le reste de ce chapitre se concentre sur ADO, en partant du principe que vous maîtrisez déjà DAO (chapitre 9) et que vous savez désormais *quand* ADO est le bon choix. La prochaine étape consiste tout naturellement à apprendre à établir une connexion — ce qui commence par la maîtrise des chaînes de connexion.


⏭️ [10.2. Chaînes de connexion pour Access (Provider=Microsoft.ACE.OLEDB)](/10-ado-access/02-chaines-connexion-access.md)
