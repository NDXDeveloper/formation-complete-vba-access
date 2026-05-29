🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 11.3. Requêtes action : INSERT, UPDATE, DELETE

La section 11.1 a expliqué **comment exécuter** une requête action depuis VBA — `CurrentDb.Execute` avec `dbFailOnError`, ou `DoCmd.RunSQL` avec précaution. Cette section s'intéresse aux requêtes action **elles-mêmes** : leur syntaxe, leurs usages et leurs pièges. Trois instructions composent cette famille : **`INSERT`** (ajouter), **`UPDATE`** (modifier) et **`DELETE`** (supprimer).

Leur force tient à leur nature **ensembliste** : une seule instruction agit sur un ensemble de lignes répondant à un critère. C'est ce qui les rend, comme nous l'avons souligné dès la section 10.5, radicalement plus rapides qu'une boucle sur un recordset pour traiter des volumes importants. Mais cette puissance a un revers : une requête action mal formulée peut affecter bien plus de lignes que prévu. La vigilance sur la clause `WHERE` sera donc le fil conducteur de cette section.

---

## INSERT : ajouter des enregistrements

L'instruction `INSERT INTO` se décline en deux formes aux usages distincts.

### Forme 1 — `INSERT ... VALUES` : ajouter une ligne

La première forme ajoute **un seul enregistrement** à partir de valeurs littérales :

```sql
INSERT INTO Clients (NomClient, Ville, Actif)
VALUES ('Dupont', 'Paris', True);
```

La liste des colonnes entre parenthèses est facultative, mais **vivement recommandée** : la nommer explicitement protège votre code si la structure de la table évolue (ajout d'une colonne, changement d'ordre). On omet généralement les colonnes **NuméroAuto**, qu'Access renseigne automatiquement.

### Forme 2 — `INSERT ... SELECT` : ajouter en masse

La seconde forme, dite *requête d'ajout*, insère **plusieurs enregistrements** issus d'une autre requête — sans clause `VALUES` :

```sql
INSERT INTO ClientsArchive (IDClient, NomClient)
SELECT IDClient, NomClient
FROM Clients
WHERE Inactif = True;
```

C'est la forme ensembliste par excellence : elle copie ou archive des centaines de lignes en une instruction, là où une boucle `AddNew` serait lente et verbeuse. Exécutée depuis VBA, cela donne :

```vba
Dim db As DAO.Database
Set db = CurrentDb
db.Execute "INSERT INTO ClientsArchive (IDClient, NomClient) " & _
           "SELECT IDClient, NomClient FROM Clients WHERE Inactif = True;", dbFailOnError
Debug.Print db.RecordsAffected & " enregistrement(s) archivé(s)."
Set db = Nothing
```

---

## UPDATE : modifier des enregistrements

L'instruction `UPDATE` modifie les valeurs d'enregistrements existants. Sa structure repose sur une clause `SET` (les colonnes à modifier) et une clause `WHERE` (les lignes concernées) :

```sql
UPDATE Clients
SET Actif = False, DateMaj = Date()
WHERE Region = 'Nord';
```

Plusieurs colonnes se modifient en une fois, séparées par des virgules. Les expressions de la clause `SET` peuvent **référencer d'autres colonnes**, ce qui permet des calculs élégants :

```sql
UPDATE Produits
SET Prix = Prix * 1.1
WHERE Categorie = 'Premium';
```

Le moteur Jet autorise par ailleurs des **`UPDATE` avec jointures**, pour modifier une table en fonction d'une autre :

```sql
UPDATE Commandes INNER JOIN Clients
ON Commandes.IDClient = Clients.IDClient
SET Commandes.Remise = 0.05
WHERE Clients.Statut = 'Fidèle';
```

(Cette syntaxe est propre à Jet ; sa correspondance en T-SQL diffère, comme le détaille l'annexe F.)

---

## DELETE : supprimer des enregistrements

L'instruction `DELETE` retire des enregistrements entiers répondant à un critère. Access admet deux écritures équivalentes, avec ou sans l'astérisque :

```sql
DELETE FROM Clients WHERE Inactif = True;
DELETE *    FROM Clients WHERE Inactif = True;   -- l'astérisque est idiomatique en Jet
```

Comme pour `UPDATE`, le moteur Jet permet une **suppression avec jointure**, en précisant la table dont les lignes doivent être supprimées :

```sql
DELETE Commandes.*
FROM Commandes INNER JOIN Clients
ON Commandes.IDClient = Clients.IDClient
WHERE Clients.Pays = 'Inconnu';
```

Une distinction importante : `DELETE` supprime des **lignes entières**. Pour seulement *vider* la valeur de certains champs sans supprimer l'enregistrement, on utilise au contraire un `UPDATE` affectant `Null` :

```sql
UPDATE Clients SET Telephone = Null WHERE Telephone = '0000000000';
```

---

## La clause `WHERE` : le garde-fou vital

Voici le point le plus critique de toute cette section. Pour `UPDATE` comme pour `DELETE`, **l'absence de clause `WHERE` applique l'opération à la table entière**.

```sql
UPDATE Clients SET Actif = False;   -- ⚠️ DÉSACTIVE TOUS LES CLIENTS
DELETE FROM Clients;                -- ⚠️ SUPPRIME TOUS LES CLIENTS
```

> ⚠️ **L'erreur catastrophique** : oublier la clause `WHERE`. Ce n'est pas une erreur de syntaxe — la requête est parfaitement valide — mais une erreur de portée aux conséquences potentiellement irréversibles. Une seconde d'inattention peut vider une table de production.

Trois habitudes protègent contre ce risque.

**Prévisualiser avec un `SELECT`.** Avant de lancer un `UPDATE` ou un `DELETE`, exécutez d'abord un `SELECT` portant **exactement la même clause `WHERE`**, pour vérifier *quelles* lignes seront affectées :

```sql
-- 1. On vérifie d'abord ce qui correspond
SELECT * FROM Clients WHERE Region = 'Nord';
-- 2. Une fois certain du périmètre, on agit
UPDATE Clients SET Actif = False WHERE Region = 'Nord';
```

**Encadrer par une transaction.** En enveloppant l'opération dans une transaction (chapitre 14), on se ménage la possibilité d'annuler (`Rollback`) si le résultat n'est pas celui attendu.

**Activer `dbFailOnError`.** Comme vu à la section 11.1, cette option garantit qu'une erreur en cours d'exécution est bien signalée plutôt qu'ignorée silencieusement.

---

## Vérifier la portée : `RecordsAffected`

Après l'exécution, la propriété `RecordsAffected` indique le **nombre exact de lignes traitées** — à condition, rappelons-le (section 11.1), de **conserver la référence** à l'objet `Database` :

```vba
Dim db As DAO.Database
Set db = CurrentDb
db.Execute "DELETE FROM Clients WHERE Inactif = True;", dbFailOnError

If db.RecordsAffected = 0 Then
    MsgBox "Aucun enregistrement supprimé : vérifiez le critère."
Else
    Debug.Print db.RecordsAffected & " enregistrement(s) supprimé(s)."
End If
Set db = Nothing
```

Cette vérification est précieuse : un `RecordsAffected` à zéro révèle un critère qui ne correspond à rien, tandis qu'un nombre anormalement élevé alerte sur une clause `WHERE` trop large.

---

## Action SQL ou boucle recordset ?

Le rappel mérite d'être répété, car c'est un réflexe de performance fondamental : pour modifier ou supprimer un ensemble de lignes selon un critère, **une requête action est bien plus rapide qu'une boucle sur un recordset**. Chaque tour de boucle implique un dialogue avec le moteur ; une requête action, elle, est traitée en une seule passe ensembliste. On réserve donc le recordset au traitement ligne à ligne assorti d'une logique conditionnelle complexe, et l'on confie les opérations de masse au SQL action (sections 10.5 et chapitre 18).

---

## Pièges courants

Au-delà de la clause `WHERE` manquante, plusieurs écueils guettent les requêtes action.

Le **formatage des valeurs littérales** est sans doute le plus fréquent dès que l'on construit le SQL dynamiquement : une date doit figurer entre dièses au format US/ISO, le séparateur décimal doit être le point, et les apostrophes des chaînes doivent être doublées. C'est l'objet entier de la section 11.7, particulièrement critique pour les `INSERT` et `UPDATE` contenant des valeurs.

Les **violations de contraintes** viennent ensuite : une insertion en doublon sur une clé unique, ou une valeur enfreignant une règle de validation, échoue. Avec `dbFailOnError`, l'erreur est levée et interceptable ; sans elle, les lignes fautives sont ignorées en silence.

L'**intégrité référentielle** concerne surtout `DELETE` : supprimer un enregistrement parent dont dépendent des enregistrements enfants échoue, sauf si une suppression en cascade est configurée (chapitre 14.6).

Enfin, quelques détails techniques : ne pas tenter d'insérer dans une colonne **NuméroAuto** ; encadrer de **crochets** `[ ]` les noms de champs comportant des espaces ou correspondant à des mots réservés (`[Date Commande]`) ; et veiller à la **cohérence des types** entre valeurs et colonnes.

---

## Valeurs variables : paramétrez

Dès qu'une requête action intègre des **valeurs d'origine externe** — une saisie utilisateur, un contenu de formulaire —, la concaténation directe dans la chaîne SQL expose aux injections et aux erreurs de formatage. La bonne pratique est de **paramétrer** la requête (section 11.5, et objet `Command` d'ADO en section 10.6) plutôt que d'y insérer les valeurs en clair. Ce principe, transversal à tout le chapitre, prend toute son importance avec les requêtes qui écrivent dans la base.

---

## Aparté : `SELECT ... INTO`, la requête création de table

Il existe une quatrième requête action, hybride : **`SELECT ... INTO`**, ou *requête de création de table*, qui crée une **nouvelle table** à partir du résultat d'un `SELECT` :

```sql
SELECT IDClient, NomClient INTO ClientsNord
FROM Clients
WHERE Region = 'Nord';
```

À la différence d'un `INSERT ... SELECT` (qui alimente une table *existante*), elle **crée** la table de destination. Elle se distingue aussi du `CREATE TABLE` du DDL (section 11.4) : `SELECT ... INTO` génère la structure *et* les données en une fois, tandis que `CREATE TABLE` ne crée qu'une structure vide. Pratique pour des tables temporaires ou des extractions, elle est à manier avec prudence, car elle écrase la table cible si celle-ci existe déjà.

---

## En résumé

Les trois requêtes action — `INSERT`, `UPDATE`, `DELETE` — modifient les données de manière **ensembliste** et bien plus efficacement qu'une boucle recordset. `INSERT` ajoute, soit une ligne via `VALUES`, soit un ensemble via `INSERT ... SELECT`. `UPDATE` modifie des colonnes selon un critère, éventuellement avec jointures et expressions calculées. `DELETE` supprime des lignes entières. Le point de vigilance absolu est la clause **`WHERE`** : son absence affecte toute la table, d'où les habitudes de prévisualisation par `SELECT`, de transaction et de `dbFailOnError`. On vérifie la portée via `RecordsAffected` (référence conservée), on se méfie du formatage des littéraux (section 11.7) et des contraintes d'intégrité (chapitre 14.6), et l'on paramètre dès que des valeurs externes entrent en jeu.

`INSERT`, `UPDATE` et `DELETE` agissent sur les **données**. Mais le SQL permet aussi d'agir sur la **structure** de la base elle-même — créer des tables, ajouter des colonnes, supprimer des objets. C'est le rôle du DDL, qu'aborde la section suivante.


⏭️ [11.4. Requêtes DDL : CREATE TABLE, ALTER TABLE, DROP TABLE](/11-sql-access-vba/04-requetes-ddl.md)
