🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 12.8. Propriétés personnalisées des objets base de données

Il nous reste un dernier aspect des objets structurels, évoqué à plusieurs reprises sans être détaillé : les **propriétés**. La section 12.4 signalait que des attributs comme `Caption` ou `Format` résidaient dans la collection `Properties` d'un champ ; la section 12.5 montrait qu'il fallait parfois les *créer* avant de les définir. Cette section élucide le mécanisme et l'étend : tout objet DAO possède une collection `Properties`, et l'on peut y stocker non seulement les propriétés reconnues par Access, mais aussi ses **propres métadonnées**.

C'est la dernière brique du modèle objet structurel — et une fonctionnalité élégante pour étendre une application sans alourdir son schéma.

---

## La collection Properties : deux natures de propriétés

Chaque objet DAO — `Database`, `TableDef`, `Field`, `QueryDef`, `Index`, `Relation` — possède une collection **`Properties`** rassemblant des objets `Property`. Ces propriétés sont de **deux natures** distinctes.

Les propriétés **intrinsèques** sont prédéfinies par DAO et **toujours présentes** : `Field.Name`, `Field.Type`, etc. On y accède directement comme membres de l'objet, ou via la collection.

Les propriétés **personnalisées** (ou définies par l'utilisateur) ne sont **pas prédéfinies** : elles n'existent qu'une fois **créées et ajoutées** à la collection. C'est le cas, vu de DAO, de propriétés pourtant familières comme `Caption` ou `Format` — Access les reconnaît, mais elles n'existent sur un champ que si on les y a ajoutées.

> 📌 **Le point crucial** : une propriété **non créée n'existe pas** dans la collection. Tenter de la **lire ou de la définir** déclenche l'erreur 3270, *« Propriété introuvable »*. C'est précisément ce comportement qui imposait, à la section 12.5, une logique « créer si absente ». Toute manipulation de propriété personnalisée doit en tenir compte.

---

## L'objet Property

Un objet `Property` se caractérise par trois attributs : `Name` (son nom), `Type` (son type de donnée, via une constante `dbText`, `dbInteger`, `dbBoolean`…) et `Value` (sa valeur). On le crée par la méthode **`CreateProperty`** de l'objet hôte, puis on l'**ajoute** à la collection :

```vba
obj.Properties.Append obj.CreateProperty(nom, type, valeur)
```

On retrouve ici le couple création/`Append` caractéristique des objets structurels DAO (sections 12.5 et 12.6).

---

## Lire une propriété en toute sécurité

Puisque l'accès à une propriété inexistante échoue, toute lecture doit être **protégée** par une gestion d'erreurs renvoyant une valeur par défaut :

```vba
Public Function LirePropriete(ByVal obj As Object, ByVal nom As String, _
                              Optional ByVal defaut As Variant = Null) As Variant
    On Error Resume Next
    LirePropriete = obj.Properties(nom).Value
    If Err.Number <> 0 Then LirePropriete = defaut
    On Error GoTo 0
End Function
```

Cette fonction renvoie la valeur de la propriété si elle existe, ou la valeur par défaut sinon — sans jamais lever d'erreur à l'appelant.

---

## Définir ou créer une propriété

Le patron complémentaire, généralisation de celui de la section 12.5, **définit** la propriété si elle existe ou la **crée** sinon. C'est l'outil de référence pour manipuler toute propriété personnalisée :

```vba
Public Sub DefinirPropriete(ByVal obj As Object, ByVal nom As String, _
                            ByVal typ As Integer, ByVal valeur As Variant)
    On Error Resume Next
    obj.Properties(nom).Value = valeur                 ' si elle existe déjà
    If Err.Number <> 0 Then                            ' sinon, on la crée
        Err.Clear
        obj.Properties.Append obj.CreateProperty(nom, typ, valeur)
    End If
    On Error GoTo 0
End Sub
```

Ces deux fonctions — lire et définir — couvrent l'essentiel des besoins, sur n'importe quel objet DAO.

---

## Les propriétés reconnues par Access

Le premier usage est la configuration des **propriétés que comprend Access**. Au niveau d'un **champ**, on retrouve `Caption` (légende), `Format` (format d'affichage), `Description`, `DecimalPlaces` (décimales) ou `InputMask` (masque de saisie). Les définir par code revient à configurer la façon dont Access affiche et traite le champ — ce que le DDL ne permet pas (section 11.4), et qui constitue l'atout de l'approche DAO (section 12.5).

```vba
Dim td As DAO.TableDef
Set td = CurrentDb.TableDefs("Produits")
DefinirPropriete td.Fields("Designation"), "Caption", dbText, "Désignation du produit"
DefinirPropriete td.Fields("Prix"), "Format", dbText, "Standard"
DefinirPropriete td.Fields("Prix"), "DecimalPlaces", dbByte, 2
```

---

## Application concrète : les options de démarrage

Une application particulièrement utile illustre la portée du mécanisme : les **options de démarrage** d'une application Access sont stockées comme **propriétés de l'objet `Database`**. Le titre de l'application (`AppTitle`), le formulaire de démarrage (`StartupForm`), ou encore l'autorisation de la touche de contournement Maj (`AllowBypassKey`) sont des propriétés personnalisées de `CurrentDb` — qui n'existent que si on les a créées.

```vba
' Définir des options de démarrage par code
DefinirPropriete CurrentDb, "AppTitle", dbText, "Gestion Commerciale"
DefinirPropriete CurrentDb, "StartupForm", dbText, "frmAccueil"
DefinirPropriete CurrentDb, "AllowBypassKey", dbBoolean, False
```

C'est exactement ainsi qu'on paramètre par code le comportement au démarrage d'une application — un sujet approfondi au chapitre 17.7, qui s'appuie sur le mécanisme décrit ici.

---

## Stocker ses propres métadonnées

Au-delà des propriétés reconnues par Access, rien n'empêche de stocker ses **métadonnées applicatives arbitraires** sur n'importe quel objet de la base. C'est un moyen léger de conserver une information sans créer de table dédiée : un numéro de version sur la `Database`, une date de dernière synchronisation sur une `TableDef`, une étiquette de catégorie sur un champ.

```vba
' Stocker la version de l'application dans la base elle-même
DefinirPropriete CurrentDb, "VersionApplication", dbText, "2.3.1"

' La relire plus tard
Debug.Print LirePropriete(CurrentDb, "VersionApplication", "inconnue")
```

Ces propriétés **persistent dans le fichier** et se relisent à volonté. Pour une **métadonnée scalaire unique** — version, configuration globale, indicateur —, c'est une solution élégante. En revanche, dès qu'il s'agit de données **multiples ou interrogeables**, une véritable **table de configuration** reste préférable : les propriétés personnalisées ne se requêtent pas comme des enregistrements.

---

## Supprimer une propriété

Une propriété personnalisée se supprime par la méthode `Delete` de la collection :

```vba
CurrentDb.Properties.Delete "VersionApplication"
```

---

## Au-delà : Containers et Documents

Pour être complet, signalons que DAO expose aussi des collections **`Containers`** et **`Documents`**, qui donnent accès aux métadonnées d'objets comme les formulaires et les états, sur lesquels on peut également stocker des propriétés. Ce mécanisme, plus avancé et d'usage moins courant, dépasse le cadre de cette section ; retenez simplement qu'il prolonge la même logique de propriétés à d'autres objets de la base.

---

## Précautions

Quelques précautions ferment cette section. Toute lecture d'une propriété potentiellement absente doit passer par une **lecture protégée** (l'erreur 3270 guette). Le **type** déclaré à la création doit correspondre à la valeur. Les propriétés personnalisées **font partie du fichier** : ce sont des éléments durables, à considérer comme une extension du schéma. Enfin, on **n'abuse pas** de ce stockage : pour des données structurées ou nombreuses, une table dédiée demeure le bon choix.

---

## En résumé

Tout objet DAO possède une collection **`Properties`** mêlant propriétés **intrinsèques** (toujours présentes) et **personnalisées** (à créer avant usage). Le point déterminant : accéder à une propriété inexistante **déclenche une erreur** (3270), d'où la nécessité d'une **lecture protégée** et d'un patron **« définir ou créer »** (`CreateProperty` + `Append`). Ce mécanisme sert d'abord à configurer les **propriétés reconnues par Access** (`Caption`, `Format`, `Description`… — l'atout de DAO sur le DDL), mais aussi à définir les **options de démarrage** (`AppTitle`, `StartupForm`, `AllowBypassKey`, propriétés de la `Database`, chapitre 17.7) et à stocker ses **propres métadonnées** (version, configuration) — une solution élégante pour une donnée scalaire unique, à réserver toutefois aux cas où une table de configuration ne s'impose pas.

---

## Conclusion du chapitre 12

Ce chapitre a complété le portrait de DAO en explorant ses **objets structurels**. Nous avons commencé par les **requêtes sauvegardées** : accès et inspection de la collection `QueryDefs` (12.1), création, modification et suppression (12.2), et paramétrage par code (12.3). Nous sommes ensuite passés aux **tables** : inspection de la structure via `TableDefs` (12.4) et création de tables, champs, clés et index (12.5). Puis aux **relations** entre tables et à l'intégrité référentielle (12.6), aux **tables liées** et à leur reliaison (12.7), et enfin aux **propriétés personnalisées** des objets (12.8).

La leçon d'ensemble rejoint celle du chapitre 9 : **DAO a deux visages**, l'accès aux *données* (chapitre 9) et la maîtrise de la *structure* (ce chapitre). Ensemble, ils font de DAO l'outil complet de manipulation d'une base Access par code — pour lire et écrire son contenu comme pour façonner et inspecter son ossature. On aura aussi noté un fil récurrent : le couple **création/`Append`** des objets structurels, et la gestion soigneuse des **erreurs** qui accompagne chaque manipulation.

Cette omniprésence de la gestion d'erreurs n'est pas un hasard. Manipuler données, SQL et structure, c'est s'exposer à de multiples sources d'échec — fichiers introuvables, contraintes violées, conflits, propriétés absentes. Savoir intercepter, diagnostiquer et traiter ces erreurs de façon robuste est une compétence transversale indispensable. C'est précisément l'objet du chapitre suivant.


⏭️ [13. Gestion des erreurs](/13-gestion-erreurs/README.md)
