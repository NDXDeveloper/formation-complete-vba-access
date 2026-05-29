🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 2.5. Références et bibliothèques COM à activer (DAO, ADO, Scripting...)

L'introduction du chapitre l'a annoncé : les références et la liaison sont des sources d'erreurs fréquentes, y compris chez les développeurs aguerris. Cette section traite le premier volet — les **références aux bibliothèques COM** — qui conditionnent les capacités de votre code. Bien les comprendre évite deux désagréments majeurs : des fonctionnalités indisponibles faute de bibliothèque activée, et des applications qui « cassent » au déploiement à cause d'une référence manquante.

## Qu'est-ce qu'une référence ?

VBA ne se limite pas à son propre langage : il peut piloter des objets fournis par d'autres composants — le moteur de base de données, Excel, Word, le système de fichiers, etc. Ces composants exposent leurs objets via une technologie nommée **COM** (*Component Object Model*), accompagnée d'une **bibliothèque de types** qui en décrit le contenu (objets, propriétés, méthodes, constantes).

Pour utiliser pleinement une telle bibliothèque, on lui ajoute une **référence** dans le projet VBA. Cette référence rend ses objets et ses constantes directement disponibles, avec l'aide à la saisie et la vérification à la compilation : c'est ce qu'on appelle la **liaison précoce** (*early binding*), détaillée à la section suivante (2.6). Une fois une bibliothèque référencée, on peut d'ailleurs explorer tout son contenu dans l'**explorateur d'objets** (<kbd>F2</kbd>, voir section 2.1).

## La boîte de dialogue Références

Les références se gèrent dans l'éditeur via **Outils > Références**. La boîte de dialogue affiche la **liste des bibliothèques disponibles** sur la machine, sous forme de cases à cocher :

- une case **cochée** signifie que la bibliothèque est active dans le projet ;
- pour en **ajouter** une, on la coche puis on valide ; si elle ne figure pas dans la liste, le bouton **Parcourir** permet de pointer son fichier (`.dll`, `.tlb`, `.olb`) ;
- le bas de la fenêtre indique l'**emplacement** du fichier sélectionné.

Un détail compte : l'**ordre** (priorité) des références, modifiable avec les flèches **Priorité**. Lorsque deux bibliothèques définissent un nom identique, c'est celle qui est **la plus haute** dans la liste qui l'emporte. Ce point, anodin en apparence, est au cœur du piège DAO/ADO décrit plus loin.

## Les bibliothèques activées par défaut

Une base `.accdb` neuve référence déjà plusieurs bibliothèques essentielles :

- **Visual Basic For Applications** — le langage lui-même ;
- **Microsoft Access XX.0 Object Library** — le modèle objet d'Access (chapitre 4) ;
- **Microsoft Office XX.0 Access Database Engine Object Library** — c'est **DAO** (moteur ACE) ;
- **OLE Automation** — services COM de base.

Conséquence pratique : **DAO fonctionne immédiatement**, sans rien ajouter. En revanche, **ADO** et **Scripting** ne sont **pas** activés par défaut : il faut les cocher explicitement.

## Les bibliothèques courantes à connaître

### DAO — l'accès aux données natif

**DAO** (*Data Access Objects*) est la bibliothèque d'accès aux données native d'Access, référencée par défaut. Elle expose `DBEngine`, `Workspace`, `Database`, `Recordset`, `TableDef`, `QueryDef`, etc. Son nom a évolué : pour le moteur ACE des `.accdb`, c'est *Microsoft Office XX.0 Access Database Engine Object Library* ; pour l'ancien moteur Jet des `.mdb`, c'était *Microsoft DAO 3.6 Object Library*. DAO est traité en profondeur au **chapitre 9**.

### ADO — l'accès générique, notamment externe

**ADO** (*ActiveX Data Objects*) est une technologie d'accès aux données plus générale, particulièrement utile pour les **bases externes** (SQL Server, Oracle…). Elle n'est **pas** activée par défaut : on coche *Microsoft ActiveX Data Objects 6.1 Library* (ou une version proche). Elle expose `Connection`, `Recordset`, `Command`, et son espace de noms est **`ADODB`**. ADO fait l'objet du **chapitre 10** ; le choix entre DAO et ADO est discuté en section 10.1.

### Scripting — fichiers et dictionnaires

**Microsoft Scripting Runtime** fournit deux objets très utilisés : le **`FileSystemObject`**, pour manipuler fichiers et dossiers (section 22.10), et le **`Dictionary`**, une collection clé/valeur performante (section 3.7). C'est l'une des premières bibliothèques que l'on ajoute couramment.

### Autres bibliothèques utiles

Selon les besoins, on rencontre aussi :

- **Microsoft Office XX.0 Object Library** — `CommandBars`, `FileDialog` (chapitre 17) ;
- **Microsoft Excel / Word / Outlook XX.0 Object Library** — pour l'automation de ces applications en liaison précoce (chapitre 22) ;
- **Microsoft XML (MSXML2)** et **Microsoft WinHTTP Services** — pour appeler des services web (section 22.7) ;
- **Microsoft VBScript Regular Expressions** — pour les expressions régulières.

## Le piège classique : le conflit DAO / ADO

Voici l'erreur la plus connue liée aux références. DAO **et** ADO définissent tous deux un type **`Recordset`** (ainsi que `Field`, et d'autres). Si les deux bibliothèques sont référencées, une déclaration **non qualifiée** devient **ambiguë** : VBA la résout en faveur de la bibliothèque la plus haute dans l'ordre des références — pas forcément celle que vous attendiez.

```vba
' Ambigu si DAO et ADO sont tous deux référencés
Dim rs As Recordset            ' ⚠️ dépend de l'ordre des références

' Sans ambiguïté : on qualifie systématiquement le type
Dim rsDao As DAO.Recordset     ' Recordset DAO
Dim rsAdo As ADODB.Recordset   ' Recordset ADO
```

La règle est simple et **impérative** : dès que DAO et ADO cohabitent, **qualifiez toujours** vos types avec leur espace de noms (`DAO.` ou `ADODB.`). C'est une bonne habitude à prendre en toutes circonstances, même quand une seule des deux est référencée — votre code reste ainsi explicite et à l'épreuve d'un ajout ultérieur de l'autre bibliothèque.

## Les références manquantes (MISSING)

Second piège, plus pernicieux encore. Si une bibliothèque référencée est **absente** de la machine — version d'Office différente, composant non installé —, la référence apparaît préfixée de **« MANQUANT : »** (*MISSING:*) dans la boîte de dialogue.

Le problème dépasse largement le code qui utilisait cette bibliothèque : une référence manquante peut **empêcher tout le projet de compiler**, et donc faire échouer du code parfaitement sain. On voit alors surgir des erreurs déroutantes — par exemple « Projet ou bibliothèque introuvable » sur des fonctions intrinsèques aussi banales que `Left`, `Date` ou `Trim` — simplement parce que le compilateur est désorienté.

C'est l'une des principales causes de pannes au **déploiement** : les numéros de version dans les noms de bibliothèques (16.0, 6.1…) traduisent des versions précises, qui ne correspondent pas toujours d'un poste à l'autre. Les remèdes :

- décocher la référence manquante, puis cocher la version correcte présente sur le poste ;
- ou, pour les bibliothèques susceptibles de varier selon les machines, recourir à la **liaison tardive**, qui se passe de référence (section 2.6).

La compatibilité entre versions et ces problèmes de déploiement sont approfondis au **chapitre 21**.

## Référence et liaison : un lien direct avec la section suivante

Tout ce qui précède concerne la **liaison précoce** : on référence une bibliothèque pour bénéficier de l'aide à la saisie, des constantes nommées et du contrôle à la compilation. À l'inverse, la **liaison tardive** (*late binding*) crée les objets sans aucune référence, au prix de ces avantages mais avec une bien meilleure portabilité face aux différences de versions. Ce compromis — quand activer une référence, quand s'en passer — est précisément l'objet de la section 2.6.

## Bonnes pratiques

- N'activez **que les références nécessaires** : chaque référence ajoute un coût de chargement et un point de défaillance potentiel.
- **Qualifiez** systématiquement les types ambigus (`DAO.Recordset`, `ADODB.Recordset`).
- Conservez **DAO**, natif et présent par défaut, pour l'accès aux données local.
- Pour les bibliothèques **variables selon les postes** (applications Office, parfois ADO), envisagez la **liaison tardive** afin de fiabiliser le déploiement.
- Donnez à votre projet un **nom explicite et unique** (section 2.2) : cela évite les conflits lorsqu'une base en référence une autre.

## À retenir

- Une **référence** (Outils > Références) rend disponible une **bibliothèque COM** et active la **liaison précoce** (aide à la saisie, constantes, contrôle à la compilation).
- Une base `.accdb` neuve référence déjà **DAO** (qui fonctionne donc d'emblée) ; **ADO** et **Scripting** doivent être **ajoutés** explicitement.
- **Piège n°1 — conflit DAO/ADO** : les deux définissent `Recordset` ; **qualifiez toujours** (`DAO.Recordset`, `ADODB.Recordset`).
- **Piège n°2 — référence MANQUANTE** : elle peut **casser tout le projet**, avec des erreurs trompeuses sur des fonctions de base ; cause fréquente de pannes au déploiement (chapitre 21).
- Pour les bibliothèques dont la version varie selon les machines, la **liaison tardive** (section 2.6) évite le recours aux références.

---


⏭️ [2.6. Liaisons précoces (Early Binding) vs liaisons tardives (Late Binding)](/02-interface-environnement/06-early-vs-late-binding.md)
