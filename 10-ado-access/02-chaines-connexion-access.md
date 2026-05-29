🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 10.2. Chaînes de connexion pour Access (`Provider=Microsoft.ACE.OLEDB`)

Toute connexion ADO commence par une **chaîne de connexion** (*connection string*). C'est le texte qui indique à ADO — et, en arrière-plan, à OLE DB — *quel* fournisseur utiliser et *où* trouver les données. Maîtriser cette chaîne est donc le premier geste technique concret du chapitre : sans connexion correctement formée, aucune des opérations des sections suivantes ne pourra fonctionner.

Cette section se concentre sur les connexions vers des **bases Access** (le moteur ACE). Les chaînes de connexion vers des serveurs externes (SQL Server, Oracle, PostgreSQL, MySQL) seront traitées en section 10.9 et récapitulées dans l'annexe D.

---

## Anatomie d'une chaîne de connexion

Une chaîne de connexion est une suite de paires **`clé=valeur`** séparées par des points-virgules :

```
Clé1=Valeur1;Clé2=Valeur2;Clé3=Valeur3;
```

L'ordre des paires n'a pas d'importance, et le point-virgule final est facultatif. Pour une base Access, deux clés sont indispensables :

La clé **`Provider`** désigne le fournisseur OLE DB qui saura dialoguer avec le moteur de la base. La clé **`Data Source`** indique le chemin complet du fichier à ouvrir. Toutes les autres clés (mot de passe, mode de partage, etc.) sont optionnelles et viennent affiner le comportement de la connexion.

---

## Le fournisseur ACE : `Microsoft.ACE.OLEDB.12.0`

Le fournisseur de référence pour Access est **`Microsoft.ACE.OLEDB.12.0`** (ACE = *Access Connectivity Engine*). C'est lui qui ouvre les fichiers `.accdb` modernes, mais il sait également lire les anciens `.mdb`.

Le numéro `12.0` peut surprendre : il est resté figé alors même qu'Access a connu de nombreuses versions depuis. Il s'agit d'un identifiant de fournisseur, et non du numéro de version d'Access. Sur certaines installations récentes d'Office, un fournisseur **`Microsoft.ACE.OLEDB.16.0`** est également enregistré ; en pratique, `12.0` reste le choix le plus universel et le plus portable, car il correspond au redistribuable largement déployé. Privilégiez-le, sauf raison contraire propre à votre environnement.

Il existe un fournisseur plus ancien, **`Microsoft.Jet.OLEDB.4.0`**, que vous rencontrerez dans du code hérité. Deux limitations majeures le condamnent aujourd'hui : il **ne peut pas ouvrir les fichiers `.accdb`** (uniquement les `.mdb`), et il **n'existe qu'en version 32 bits** — un point de blocage que nous détaillerons plus loin. Sauf maintenance de code ancien sur des `.mdb`, on ne l'utilise plus.

Le tableau suivant résume le bon fournisseur selon le format de fichier.

| Format de fichier | Fournisseur recommandé | Remarques |
|---|---|---|
| `.accdb` (Access 2007+) | `Microsoft.ACE.OLEDB.12.0` | Seul ACE sait ouvrir le format `.accdb` |
| `.accde` / `.accdr` | `Microsoft.ACE.OLEDB.12.0` | Identique au `.accdb` |
| `.mdb` (Access 97–2003) | `Microsoft.ACE.OLEDB.12.0` (ou `Jet.OLEDB.4.0`) | Jet 4.0 = 32 bits uniquement, code hérité |

---

## La chaîne minimale

Pour ouvrir une base `.accdb` sans protection particulière, la chaîne se réduit à l'essentiel :

```
Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Donnees\Gestion.accdb;
```

Côté code VBA, cela donne :

```vba
Dim cn As ADODB.Connection
Set cn = New ADODB.Connection
cn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Donnees\Gestion.accdb;"
' ... travail sur la connexion ...
cn.Close
Set cn = Nothing
```

C'est fonctionnel, mais coder le chemin en dur est rarement une bonne idée — surtout quand la base que l'on veut interroger est… celle dans laquelle le code s'exécute déjà. D'où le raccourci décisif de la section suivante.

---

## Le raccourci à connaître : `CurrentProject.Connection`

Lorsqu'on souhaite manipuler **la base Access courante** en ADO, il est inutile — et même contre-productif — de reconstruire une chaîne de connexion à la main. Access expose déjà une connexion ADO toute prête vers la base active :

```vba
Dim cn As ADODB.Connection
Set cn = CurrentProject.Connection   ' connexion ADO déjà ouverte sur la base courante

Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
rs.Open "SELECT * FROM Clients", cn, adOpenForwardOnly, adLockReadOnly
' ...
rs.Close
Set rs = Nothing
' Ne PAS fermer cn : c'est la connexion d'Access lui-même
```

`CurrentProject.Connection` renvoie un objet `Connection` **déjà ouvert** sur la base en cours. Avantage immédiat : aucun chemin codé en dur, donc aucune adaptation lorsque la base est déplacée ou renommée. Un point de vigilance important : comme cette connexion appartient à Access, **on ne la ferme pas** soi-même — on se contente de l'utiliser.

Si vous avez besoin de la *chaîne* elle-même (par exemple pour la transmettre ailleurs ou la déboguer), la propriété **`CurrentProject.BaseConnectionString`** la renvoie sous forme de texte :

```vba
Debug.Print CurrentProject.BaseConnectionString
' Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\...\Gestion.accdb;
'   Mode=Share Deny None;Jet OLEDB:Engine Type=6;...
```

Les versions récentes d'Access exposent aussi `CurrentProject.AccessConnection`, variante qui force l'usage du fournisseur ACE. Dans la majorité des cas, `CurrentProject.Connection` suffit amplement.

> 💡 **À retenir** : pour la base courante, partez systématiquement de `CurrentProject.Connection`. Ne construisez une chaîne manuellement que pour atteindre une base *autre* que celle en cours d'exécution.

---

## Les paramètres courants

Au-delà du couple `Provider` / `Data Source`, quelques paramètres reviennent régulièrement. Le tableau ci-dessous les récapitule, avant que nous ne détaillions les plus importants.

| Paramètre | Rôle | Exemple de valeur |
|---|---|---|
| `Provider` | Fournisseur OLE DB | `Microsoft.ACE.OLEDB.12.0` |
| `Data Source` | Chemin du fichier de base | `C:\Donnees\Gestion.accdb` |
| `Jet OLEDB:Database Password` | Mot de passe de la base | `M0nM0tDeP@sse` |
| `Mode` | Mode de partage / verrouillage | `Share Deny None` |
| `Persist Security Info` | Conservation du mot de passe en mémoire | `False` |
| `Extended Properties` | Options pour sources non-Access (Excel, texte) | `"Excel 12.0 Xml;HDR=YES"` |

### Mot de passe de base

Si la base est protégée par un mot de passe (chapitre 20.2), il se fournit via le paramètre **`Jet OLEDB:Database Password`**. Le préfixe `Jet OLEDB:` peut sembler incongru avec le fournisseur ACE, mais il a été conservé tel quel pour des raisons de compatibilité : c'est bien le nom attendu, même en ACE.

```vba
cn.Open "Provider=Microsoft.ACE.OLEDB.12.0;" & _
        "Data Source=C:\Donnees\Gestion.accdb;" & _
        "Jet OLEDB:Database Password=M0nM0tDeP@sse;"
```

Notez que ce mécanisme correspond au mot de passe *de la base de données* (format `.accdb`). L'ancienne **sécurité au niveau utilisateur** (fichiers `.mdw`, comptes et groupes) n'existe que pour les `.mdb` avec le fournisseur Jet ; elle a été abandonnée avec le format `.accdb`. Les aspects de sécurité sont détaillés au chapitre 20.

### Mode de partage

Le paramètre **`Mode`** contrôle la manière dont la connexion partage le fichier avec les autres utilisateurs. Les valeurs les plus fréquentes sont `Share Deny None` (partage complet, le cas multi-utilisateurs courant), `Share Exclusive` (ouverture exclusive, qui interdit tout autre accès simultané) ou encore `Read` (lecture seule). Le mode de partage est un point sensible en environnement concurrent, traité plus largement au chapitre 15.

### Sécurité de l'information de connexion

Le paramètre **`Persist Security Info`** indique si les informations sensibles — typiquement le mot de passe — restent récupérables dans la propriété `ConnectionString` *après* l'ouverture. Par prudence, on le positionne à **`False`** afin que le mot de passe ne reste pas exposé en mémoire une fois la connexion établie.

---

## Le point sensible : 32 bits contre 64 bits

C'est l'une des sources d'erreurs les plus déroutantes pour qui déploie une application Access. Le message classique — *« Le fournisseur Microsoft.ACE.OLEDB.12.0 n'est pas inscrit sur l'ordinateur local »* — découle presque toujours d'un désaccord d'architecture.

Deux règles à garder en tête. D'une part, le **fournisseur Jet 4.0 n'existe qu'en 32 bits** : une application Access 64 bits ne peut tout simplement pas l'utiliser. D'autre part, le **fournisseur ACE existe en 32 et en 64 bits, mais sa version doit correspondre à celle d'Office** installé sur le poste. Un Office 32 bits exige l'ACE 32 bits, un Office 64 bits exige l'ACE 64 bits — et il est délicat d'installer les deux simultanément sur une même machine.

En pratique, la chaîne de connexion `Microsoft.ACE.OLEDB.12.0` est identique quelle que soit l'architecture ; c'est le *fournisseur installé* qui doit être cohérent avec l'application. Ce sujet, lourd de conséquences au déploiement, est approfondi au chapitre 21.7.

---

## Cas particuliers : lire de l'Excel ou du texte via ACE

Le fournisseur ACE ne se limite pas aux bases Access : il sait aussi attaquer des classeurs Excel et des fichiers texte, ce qui en fait un outil d'import léger. La spécificité tient au paramètre **`Extended Properties`**, qui précise le type de source et ses options.

Pour un classeur Excel :

```vba
cn.Open "Provider=Microsoft.ACE.OLEDB.12.0;" & _
        "Data Source=C:\Donnees\Ventes.xlsx;" & _
        "Extended Properties=""Excel 12.0 Xml;HDR=YES"";"
```

Ici, `HDR=YES` indique que la première ligne contient les en-têtes de colonnes. Pour un fichier texte délimité (CSV), `Data Source` désigne le **dossier** contenant le fichier, et `Extended Properties` précise le format :

```vba
cn.Open "Provider=Microsoft.ACE.OLEDB.12.0;" & _
        "Data Source=C:\Donnees\;" & _
        "Extended Properties=""text;HDR=YES;FMT=Delimited"";"
```

Ces usages sont pratiques pour des imports ponctuels, mais Access offre par ailleurs des méthodes dédiées (`DoCmd.TransferSpreadsheet`, `DoCmd.TransferText`, vues au chapitre 5.5) souvent plus simples lorsqu'il s'agit d'intégrer durablement les données.

---

## Bonnes pratiques de construction

Quelques principes évitent la plupart des déboires liés aux chaînes de connexion.

**Ne codez jamais un chemin en dur sans nécessité.** Pour la base courante, utilisez `CurrentProject.Connection`. Pour une base voisine, construisez le chemin relativement à `CurrentProject.Path` plutôt que de figer un chemin absolu qui se brisera au moindre déplacement.

**Centralisez la construction de la chaîne** dans une fonction unique. Plutôt que de disséminer la chaîne dans tout le projet, une fonction dédiée garantit la cohérence et facilite la maintenance — un seul endroit à modifier en cas de changement. Ce réflexe rejoint les principes d'architecture du chapitre 16.

```vba
Public Function ChaineConnexionLocale() As String
    ChaineConnexionLocale = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
                            "Data Source=" & CurrentProject.Path & _
                            "\DonneesMetier.accdb;"
End Function
```

**Attention aux espaces et aux caractères spéciaux.** Une valeur peut contenir des espaces sans guillemets (un chemin tel que `C:\Mes Documents\base.accdb` est accepté tel quel). En revanche, dès qu'une valeur contient elle-même un point-virgule — fréquent dans `Extended Properties` — il faut l'encadrer de guillemets, ce qui, en VBA, se traduit par des guillemets doublés (`""`).

**Protégez les informations sensibles.** Évitez de stocker un mot de passe en clair dans le code source ou dans une chaîne persistée, et positionnez `Persist Security Info=False`. La gestion des secrets est abordée au chapitre 20.

---

## Et pour les bases externes ?

Tout ce qui précède concerne le moteur ACE. Dès qu'il s'agit d'atteindre un **serveur SQL Server, une base Oracle, PostgreSQL ou MySQL**, la logique reste la même — on change simplement de fournisseur et l'on ajuste les paramètres (serveur, base, authentification). Ces chaînes spécifiques, plus variées, font l'objet de la section 10.9, avec un récapitulatif pratique en annexe D. Le site communautaire de référence *connectionstrings.com* constitue par ailleurs un aide-mémoire précieux pour retrouver la syntaxe exacte d'un fournisseur donné.

---

## En résumé

Une chaîne de connexion Access repose sur deux clés essentielles, `Provider` et `Data Source`, le fournisseur de référence étant `Microsoft.ACE.OLEDB.12.0`. Pour la base courante, le réflexe à adopter est `CurrentProject.Connection`, qui dispense de toute chaîne codée en dur. Les paramètres optionnels (mot de passe, mode de partage, sécurité) affinent le comportement, tandis que la question du 32/64 bits demeure le principal piège au déploiement. Enfin, une construction centralisée et prudente des chaînes épargne bien des erreurs à long terme.

La chaîne n'étant qu'un texte décrivant *comment* se connecter, l'étape suivante consiste à donner vie à cette description : ouvrir, configurer et fermer proprement l'objet qui porte la connexion.


⏭️ [10.3. Objet Connection — ouverture et fermeture](/10-ado-access/03-objet-connection.md)
